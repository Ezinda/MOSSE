VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form frmotrosparam 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Otros Parametros"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   8940
   Begin TabDlg.SSTab SSTab1 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   13150
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Generales"
      TabPicture(0)   =   "frmotrosparam.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Cuentas de fondos y Resultados"
      TabPicture(1)   =   "frmotrosparam.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label8(0)"
      Tab(1).Control(1)=   "Label8(1)"
      Tab(1).Control(2)=   "Label8(2)"
      Tab(1).Control(3)=   "Label8(3)"
      Tab(1).Control(4)=   "Label8(4)"
      Tab(1).Control(5)=   "Check4"
      Tab(1).Control(6)=   "Text2(2)"
      Tab(1).Control(7)=   "List1"
      Tab(1).Control(8)=   "DataGrid3"
      Tab(1).Control(9)=   "nuevo"
      Tab(1).Control(10)=   "grabar"
      Tab(1).Control(11)=   "Text2(1)"
      Tab(1).Control(12)=   "DataCombo2"
      Tab(1).Control(13)=   "Frame5"
      Tab(1).Control(14)=   "Command4"
      Tab(1).Control(15)=   "Text2(0)"
      Tab(1).Control(16)=   "Text2(3)"
      Tab(1).ControlCount=   17
      Begin VB.Frame Frame7 
         Caption         =   "Plan de Cuentas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   5640
         TabIndex        =   46
         Top             =   5040
         Width           =   2775
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            DataField       =   "placcuentasunif"
            DataSource      =   "datempresa1"
            Height          =   375
            Left            =   2040
            TabIndex        =   48
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label9 
            Caption         =   "Un Plan para todos los Periodos:"
            Height          =   375
            Left            =   120
            TabIndex        =   47
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Nuevo"
         Height          =   495
         Left            =   5640
         TabIndex        =   45
         Top             =   6720
         Width           =   855
      End
      Begin VB.Frame Frame6 
         Caption         =   "Instrumentos de Cobro y Pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   120
         TabIndex        =   43
         Top             =   5040
         Width           =   5415
         Begin MSDataGridLib.DataGrid DataGrid5 
            Bindings        =   "frmotrosparam.frx":0038
            Height          =   1575
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   2778
            _Version        =   393216
            AllowUpdate     =   -1  'True
            AllowArrows     =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            TabAction       =   2
            WrapCellPointer =   -1  'True
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
      End
      Begin VB.TextBox Text2 
         DataField       =   "grupo"
         DataSource      =   "datparamresultados"
         Height          =   285
         Index           =   3
         Left            =   -69360
         MaxLength       =   30
         TabIndex        =   41
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         DataField       =   "id"
         DataSource      =   "datparamresultados"
         Height          =   285
         Index           =   0
         Left            =   -69360
         MaxLength       =   30
         TabIndex        =   38
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Nulo"
         Height          =   240
         Left            =   -68280
         TabIndex        =   36
         Top             =   840
         Width           =   735
      End
      Begin VB.Frame Frame5 
         Caption         =   "Otras Cuentas y Prorrateo"
         Height          =   2055
         Left            =   -70560
         TabIndex        =   32
         Top             =   3600
         Width           =   3975
         Begin VB.CommandButton Command3 
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
            Left            =   3480
            TabIndex        =   35
            Top             =   840
            Width           =   375
         End
         Begin VB.CommandButton Command2 
            Caption         =   "G"
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
            Left            =   3480
            TabIndex        =   34
            Top             =   360
            Width           =   375
         End
         Begin MSDataGridLib.DataGrid DataGrid4 
            Bindings        =   "frmotrosparam.frx":0050
            Height          =   1695
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   2990
            _Version        =   393216
            BackColor       =   16777152
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
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "frmotrosparam.frx":0072
         Height          =   315
         Left            =   -70560
         TabIndex        =   31
         Top             =   1080
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   "codigo"
         BoundColumn     =   "idcuenta"
         Text            =   ""
      End
      Begin VB.TextBox Text2 
         DataField       =   "nombrefondo"
         DataSource      =   "datparamresultados"
         Height          =   285
         Index           =   1
         Left            =   -69360
         MaxLength       =   50
         TabIndex        =   27
         Top             =   600
         Width           =   2775
      End
      Begin VB.CommandButton grabar 
         Caption         =   "&Grabar"
         Height          =   615
         Left            =   -69240
         Picture         =   "frmotrosparam.frx":008B
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   5760
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton nuevo 
         Caption         =   "&Nuevo"
         Height          =   615
         Left            =   -70560
         Picture         =   "frmotrosparam.frx":05BD
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   5760
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "frmotrosparam.frx":0AEF
         Height          =   1455
         Left            =   -70560
         TabIndex        =   24
         Top             =   2040
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   2566
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   12648447
         HeadLines       =   1
         RowHeight       =   15
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
      Begin VB.ListBox List1 
         Height          =   5910
         ItemData        =   "frmotrosparam.frx":0B10
         Left            =   -74640
         List            =   "frmotrosparam.frx":0B12
         Style           =   1  'Checkbox
         TabIndex        =   23
         Top             =   480
         Width           =   3855
      End
      Begin VB.Frame Frame1 
         Caption         =   "Condiciones Tributarias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   4095
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "frmotrosparam.frx":0B14
            Height          =   1935
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   3413
            _Version        =   393216
            AllowArrows     =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            TabAction       =   1
            WrapCellPointer =   -1  'True
            AllowAddNew     =   -1  'True
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
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipos de Clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   4320
         TabIndex        =   18
         Top             =   600
         Width           =   4095
         Begin VB.CommandButton agregar 
            Caption         =   "&Agregar"
            Height          =   375
            Left            =   1440
            TabIndex        =   19
            Top             =   2040
            Width           =   1335
         End
         Begin MSDataGridLib.DataGrid DataGrid2 
            Bindings        =   "frmotrosparam.frx":0B2E
            Height          =   1575
            Left            =   240
            TabIndex        =   20
            Top             =   360
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   2778
            _Version        =   393216
            AllowUpdate     =   -1  'True
            AllowArrows     =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            TabAction       =   1
            WrapCellPointer =   -1  'True
            FormatLocked    =   -1  'True
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
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "codigo"
               Caption         =   "codigo"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
               DataField       =   "tipoclientes"
               Caption         =   "tipoclientes"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
               BeginProperty Column02 
                  Object.Visible         =   0   'False
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "BD Clientes y Proveedores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         TabIndex        =   15
         Top             =   3240
         Width           =   4095
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "frmotrosparam.frx":0B4C
            Height          =   315
            Left            =   360
            TabIndex        =   16
            Top             =   840
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            ListField       =   "razonsocial"
            BoundColumn     =   "empresa"
            Text            =   "DataCombo1"
         End
         Begin VB.Label Label1 
            Caption         =   "Comparte con empresa"
            Height          =   255
            Left            =   960
            TabIndex        =   17
            Top             =   480
            Width           =   1815
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Otros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   4320
         TabIndex        =   1
         Top             =   3240
         Width           =   4095
         Begin VB.CheckBox Check1 
            DataField       =   "preguntacai"
            DataSource      =   "datparamgral"
            Height          =   255
            Left            =   1800
            TabIndex        =   7
            Top             =   960
            Width           =   255
         End
         Begin VB.CheckBox Check2 
            DataField       =   "activamultiempresa"
            DataSource      =   "datparamgral"
            Height          =   255
            Left            =   1800
            TabIndex        =   6
            Top             =   1320
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Height          =   255
            Left            =   2520
            TabIndex        =   5
            Top             =   720
            Width           =   255
         End
         Begin VB.OptionButton Option2 
            Height          =   255
            Left            =   3240
            TabIndex        =   4
            Top             =   720
            Width           =   255
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Check3"
            DataField       =   "numeroautoventas"
            DataSource      =   "datparamgral"
            Height          =   255
            Left            =   2880
            TabIndex        =   3
            Top             =   1320
            Width           =   255
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Grab"
            Height          =   255
            Left            =   3360
            TabIndex        =   2
            Top             =   1320
            Width           =   615
         End
         Begin VB.PictureBox efectivo 
            DataField       =   "montomaxefectivo"
            DataSource      =   "datparamgral"
            Height          =   375
            Left            =   120
            ScaleHeight     =   315
            ScaleWidth      =   1875
            TabIndex        =   8
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label2 
            Caption         =   "Pago Máximo en efectivo"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label3 
            Caption         =   "Pregunta por CAI:"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Activa Multiempresa:"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "Ordener Cuentas"
            Height          =   255
            Left            =   2400
            TabIndex        =   11
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Codigo     Nombre"
            Height          =   255
            Left            =   2280
            TabIndex        =   10
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "Num.Auto.L.Ventas"
            Height          =   255
            Left            =   2400
            TabIndex        =   9
            Top             =   1080
            Width           =   1335
         End
      End
      Begin VB.TextBox Text2 
         DataField       =   "cuentacabecera"
         DataSource      =   "datparamresultados"
         Height          =   285
         Index           =   2
         Left            =   -69360
         MaxLength       =   30
         TabIndex        =   29
         Top             =   840
         Width           =   1095
      End
      Begin VB.CheckBox Check4 
         DataField       =   "activado"
         DataSource      =   "datparamresultados"
         Height          =   255
         Left            =   -69360
         TabIndex        =   42
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Activado:"
         Height          =   255
         Index           =   4
         Left            =   -70560
         TabIndex        =   40
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Grupo:"
         Height          =   255
         Index           =   3
         Left            =   -70560
         TabIndex        =   39
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Cod.de Fondo:"
         Height          =   255
         Index           =   2
         Left            =   -70560
         TabIndex        =   37
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Cuenta Madre:"
         Height          =   255
         Index           =   1
         Left            =   -70560
         TabIndex        =   30
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Nombre Fondo:"
         Height          =   255
         Index           =   0
         Left            =   -70560
         TabIndex        =   28
         Top             =   600
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc datacontrib 
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
   Begin MSAdodcLib.Adodc dattipoclientes 
      Height          =   330
      Left            =   1320
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
   Begin MSAdodcLib.Adodc datempresa 
      Height          =   330
      Left            =   2760
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
   Begin MSAdodcLib.Adodc datparamgral 
      Height          =   330
      Left            =   4680
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
      LcK2            =   $"frmotrosparam.frx":0B65
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
   Begin MSAdodcLib.Adodc datparamresultados 
      Height          =   330
      Left            =   7320
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
   Begin MSAdodcLib.Adodc datcuentas 
      Height          =   330
      Left            =   6000
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
   Begin MSAdodcLib.Adodc datparamresultados1 
      Height          =   330
      Left            =   3600
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
   Begin MSAdodcLib.Adodc datinstru 
      Height          =   330
      Left            =   5760
      Top             =   7440
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
   Begin MSAdodcLib.Adodc datempresa1 
      Height          =   330
      Left            =   120
      Top             =   7440
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
Attribute VB_Name = "frmotrosparam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private Sub agregar_Click()

    dattipoclientes.Recordset.AddNew
    dattipoclientes.Recordset.Fields("empresa") = login.empresaact


End Sub

Private Sub Command1_Click()

    If Option2.Value = True Then
        datparamgral.Recordset.Fields("ordenlista") = 1
    Else
       datparamgral.Recordset.Fields("ordenlista") = 0
    End If

    datparamgral.Recordset.UpdateBatch adAffectCurrent
    
    mensa = MsgBox("Debe reiniciar el Programa para que los cambios tengan efecto", vbInformation, "Atención")
    
    
    
End Sub

Private Sub Command2_Click()
On Error Resume Next
datparamresultados1.Recordset.UpdateBatch adAffectCurrent

End Sub

Private Sub Command3_Click()
On Error Resume Next
    mensa = MsgBox("Esta por borrar una cuenta de este Fondo, Esta Seguro", vbYesNo, "!! Atencion !!")
    If mensa = vbYes Then datparamresultados1.Recordset.Delete adAffectCurrent

End Sub

Private Sub Command4_Click()
    
    Text2(2).Text = ""

End Sub

Private Sub Command5_Click()

    datinstru.Recordset.AddNew
    DataGrid5.SetFocus

End Sub

Private Sub DataCombo1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        combo = Val(DataCombo1.BoundText)
        datempresa.RecordSource = "select empresa.* from empresa where empresa = " & login.empresaact & " "
        datempresa.Refresh
        
        datempresa.Recordset.Fields("bdclientes") = combo
        datempresa.Recordset.UpdateBatch adAffectCurrent
        
        datempresa.RecordSource = "select empresa.* from empresa where empresa <> " & login.empresaact & " "
        datempresa.Refresh
        
    End If

End Sub

Private Sub DataCombo2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text2(2).Text = DataCombo2.BoundText
        Text2(3).SetFocus
    End If

End Sub

Private Sub DataCombo2_KeyUp(KeyCode As Integer, Shift As Integer)
    Text2(2).Text = DataCombo2.BoundText
End Sub

Private Sub DataGrid2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And DataGrid2.Col = 1 Then
        KeyAscii = 0
        dattipoclientes.Recordset.Fields(1) = DataGrid2.Columns(1).Text
        dattipoclientes.Recordset.UpdateBatch adAffectCurrent
    End If

End Sub

Private Sub DataGrid3_Click()
On Error Resume Next
    DataCombo2.BoundText = DataGrid3.Columns(2).Text
    datparamresultados1.RecordSource = "select paramresultados1.* from paramresultados1 where empresa = " & login.empresaact & " and id = " & datparamresultados.Recordset.Fields("id") & " order by cuenta"
    datparamresultados1.Refresh

End Sub

Private Sub DataGrid3_KeyUp(KeyCode As Integer, Shift As Integer)

        DataCombo2.BoundText = DataGrid3.Columns(2).Text
        datparamresultados1.RecordSource = "select paramresultados1.* from paramresultados1 where empresa = " & login.empresaact & " and id = " & datparamresultados.Recordset.Fields("id") & " order by cuenta"
        datparamresultados1.Refresh

End Sub

Private Sub DataGrid4_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then
         KeyAscii = 0
         datparamresultados1.Recordset.UpdateBatch adAffectCurrent
    End If

End Sub

Private Sub DataGrid5_ColEdit(ByVal ColIndex As Integer)
On Error Resume Next
    If ColIndex = 1 Then
        z_cuentas.menucuentas = ""
        z_cuentas.Show
        DataGrid5.Columns(2).Text = login.empresaact
    End If

End Sub

Private Sub DataGrid5_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}", False
    End If

End Sub

Private Sub Form_Load()
On Error Resume Next
frmotrosparam.Left = 0
frmotrosparam.Top = 0

DataGrid2.Columns(0).Width = 800
DataGrid2.Columns(1).Width = 2500
DataGrid3.Columns(0).Visible = False
Rem DataGrid3.Columns(2).Visible = False
DataGrid3.Columns(1).Width = 3500

DataGrid4.Columns(0).Visible = False


datacontrib.ConnectionString = login.conexiontotal
datinstru.ConnectionString = login.conexiontotal
dattipoclientes.ConnectionString = login.conexiontotal
datempresa.ConnectionString = login.conexiontotal
datempresa1.ConnectionString = login.conexiontotal
datparamgral.ConnectionString = login.conexiontotal
datcuentas.ConnectionString = login.conexiontotal
datparamresultados.ConnectionString = login.conexiontotal
datparamresultados1.ConnectionString = login.conexiontotal
    
  datacontrib.RecordSource = "select condtrib.* from condtrib"
  datacontrib.Refresh
  
  datparamresultados.RecordSource = "select paramresultados.* from paramresultados where empresa = " & login.empresaact & " order by id"
  datparamresultados.Refresh

  datparamresultados1.RecordSource = "select paramresultados1.* from paramresultados1 where empresa = " & login.empresaact & " order by cuenta"
  datparamresultados1.Refresh

  datparamgral.RecordSource = "select parametrosgenerales.* from parametrosgenerales"
  datparamgral.Refresh
    
  dattipoclientes.RecordSource = "select tipoclientes.* from tipoclientes where empresa = " & login.empresaact & ""
  dattipoclientes.Refresh
  
  datempresa.RecordSource = "select empresa.* from empresa where empresa = " & login.empresaact & " "
  datempresa.Refresh
  
  datempresa1.RecordSource = "select empresa.* from empresa where empresa = " & login.empresaact & " "
  datempresa1.Refresh
  
  datcuentas.RecordSource = "select listacuentas.* from listacuentas WHERE empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' ORDER BY IDCUENTA"
  datcuentas.Refresh
    
  basecompartida = datempresa.Recordset.Fields("bdclientes")
    
  datempresa.RecordSource = "select empresa.* from empresa where empresa = " & basecompartida & ""
  datempresa.Refresh
  
  datinstru.RecordSource = "select instrumentospagos.* from instrumentospagos where empresa = " & login.empresaact & "  "
  datinstru.Refresh
  
DataGrid5.Columns(2).Visible = False
DataGrid5.Columns(0).Width = 3500
DataGrid5.Columns(1).Width = 1000
DataGrid5.Columns(1).Alignment = dbgCenter
  
  If datempresa.Recordset.EOF = False Then
    DataCombo1.Text = datempresa.Recordset.Fields("razonsocial")
  Else
    DataCombo1.Text = ""
  End If

  datempresa.RecordSource = "select empresa.* from empresa where empresa <> " & login.empresaact & ""
  datempresa.Refresh
    

  If datparamgral.Recordset.EOF = True Then
        datparamgral.Recordset.AddNew
        datparamgral.Recordset.Fields("montomaxefectivo") = 0
        datparamgral.Recordset.Fields("preguntacai") = 0
        datparamgral.Recordset.Fields("activamultiempresa") = 1
        datparamgral.Recordset.Fields("ordenlista") = 0
        datparamgral.Recordset.Fields("numeroautoventas") = 0
        datparamgral.Recordset.UpdateBatch adAffectCurrent
 End If
 
 If datparamgral.Recordset.Fields("ordenlista") = 0 Then
    Option1.Value = True
 Else
    Option2.Value = True
 End If
        
If datcuentas.Recordset.EOF = True Then Exit Sub

i = 0
datcuentas.Recordset.MoveFirst
Do While Not datcuentas.Recordset.EOF
      
    List1.AddItem (datcuentas.Recordset.Fields("codigo"))
    List1.Selected(i) = False
    i = i + 1
    datcuentas.Recordset.MoveNext
  
Loop
  
  datcuentas.RecordSource = "select listacuentas2.* from listacuentas2 WHERE empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' ORDER BY IDCUENTA"
  datcuentas.Refresh
    
End Sub

Private Sub grabar_Click()


    datparamresultados.Recordset.Fields("empresa") = login.empresaact
    
    
    
    datparamresultados.Recordset.UpdateBatch adAffectCurrent


End Sub

Private Sub List1_ItemCheck(Item As Integer)
On Error Resume Next
    If List1.Selected(List1.ListIndex) = True Then
        For x = 1 To 15
            guion = Mid(List1.Text, x, 1)
            If guion = " " Then GoTo sale
        Next x
sale:
        datparamresultados1.Recordset.AddNew
        datparamresultados1.Recordset.Fields("id") = datparamresultados.Recordset.Fields("id")
        datparamresultados1.Recordset.Fields("empresa") = datparamresultados.Recordset.Fields("empresa")
        datparamresultados1.Recordset.Fields("prorrateo") = 100
        datparamresultados1.Recordset.Fields("cuenta") = Val(Left(List1.Text, x - 1))
        datparamresultados1.Recordset.UpdateBatch adAffectCurrent
    End If

End Sub

Private Sub nuevo_Click()

    datparamresultados.Recordset.AddNew

End Sub

Private Sub Text1_Change()
On Error Resume Next
    datempresa1.Recordset.UpdateBatch adAffectCurrent

End Sub

Private Sub Text2_GotFocus(Index As Integer)
        
        If Index = 2 Then
            DataCombo2.SetFocus
            Exit Sub
        End If


End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        If Index = 3 Then
            Check4.SetFocus
            Exit Sub
        End If
        Text2(Index + 1).SetFocus
    End If

End Sub

