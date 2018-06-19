VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form frmajusteproveedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ajustes de EC de Proveedores"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8715
   LinkTopic       =   "Ajustes EC Clientes"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   8715
   Begin VB.CommandButton Command6 
      Caption         =   "limpia"
      Height          =   255
      Left            =   6120
      TabIndex        =   60
      Top             =   7200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   240
      TabIndex        =   13
      Top             =   960
      Width           =   8175
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2040
         TabIndex        =   15
         Text            =   "Combo1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton buscar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   315
         Left            =   3240
         TabIndex        =   16
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   13
         Mask            =   "####-########"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo Comp.:"
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   8175
      Begin VB.CommandButton Command5 
         Caption         =   "Buscar Proveedor:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Fecha Comp.:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Proveedor:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Total Comp.:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Saldo Comp.:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Comprobante"
         Height          =   255
         Index           =   6
         Left            =   3720
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Reg.Contable:"
         Height          =   255
         Index           =   7
         Left            =   4680
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Periodo:"
         Height          =   255
         Index           =   8
         Left            =   3480
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000016&
         Height          =   315
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   600
         Width           =   615
      End
      Begin MSDataListLib.DataCombo DataCombo4 
         Bindings        =   "frmajusteproveedores.frx":0000
         Height          =   315
         Left            =   1920
         TabIndex        =   32
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         BackColor       =   -2147483626
         ListField       =   "razonsocial"
         BoundColumn     =   "codproveedor"
         Text            =   ""
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   4
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   5
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   6
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   960
         Width           =   1695
      End
      Begin MSDataListLib.DataCombo DataCombo6 
         Bindings        =   "frmajusteproveedores.frx":001A
         Height          =   315
         Left            =   4920
         TabIndex        =   33
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         BackColor       =   -2147483626
         ListField       =   "numcompr"
         BoundColumn     =   "tipocompr"
         Text            =   ""
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "CUIT"
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
         Left            =   7200
         TabIndex        =   12
         Top             =   1320
         Width           =   735
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   1147
      TabCaption(0)   =   "Cambiar el Proveedor de Una Factura o NC"
      TabPicture(0)   =   "frmajusteproveedores.frx":0037
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grabar"
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(2)=   "datlibroventas"
      Tab(0).Control(3)=   "datclientes"
      Tab(0).Control(4)=   "datlibroventas1"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Reasignar Pagos mal asignados a Facturas"
      TabPicture(1)   =   "frmajusteproveedores.frx":0053
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "datasigcomp"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "datrecibos"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "datrecibosabonan"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame4"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "List1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Text1(7)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command1"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Command5(18)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Asig. Parcial o Total de Pagos S/C a Facturas"
      TabPicture(2)   =   "frmajusteproveedores.frx":006F
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "datasiento1"
      Tab(2).Control(1)=   "DataGrid4"
      Tab(2).Control(2)=   "DataGrid3"
      Tab(2).Control(3)=   "datrecibosabonan1"
      Tab(2).Control(4)=   "datrecibosinc"
      Tab(2).Control(5)=   "Frame5"
      Tab(2).Control(6)=   "Text1(10)"
      Tab(2).Control(7)=   "Command2"
      Tab(2).Control(8)=   "DataGrid1"
      Tab(2).Control(9)=   "Command5(14)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "Asignacion Parcial o total de NC a Facturas"
      TabPicture(3)   =   "frmajusteproveedores.frx":008B
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).Control(1)=   "DataGrid2"
      Tab(3).Control(2)=   "Command3"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Ajuste Facturas"
      TabPicture(4)   =   "frmajusteproveedores.frx":00A7
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Command4"
      Tab(4).Control(1)=   "Frame7"
      Tab(4).Control(2)=   "grabaasiento"
      Tab(4).Control(3)=   "datasiento"
      Tab(4).Control(4)=   "datmaestro"
      Tab(4).ControlCount=   5
      Begin VB.CommandButton Command4 
         Caption         =   "Grabar"
         Height          =   615
         Left            =   -71640
         Picture         =   "frmajusteproveedores.frx":00C3
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   5520
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin VB.Frame Frame7 
         Caption         =   "Ajuste"
         Height          =   1455
         Left            =   -74880
         TabIndex        =   61
         Top             =   3840
         Width           =   8175
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6480
            TabIndex        =   68
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1080
            TabIndex        =   67
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   4200
            TabIndex        =   66
            Top             =   480
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1080
            TabIndex        =   65
            Top             =   840
            Width           =   1335
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Debe:"
            Height          =   255
            Index           =   19
            Left            =   240
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   480
            Width           =   735
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Haber:"
            Height          =   255
            Index           =   20
            Left            =   240
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   840
            Width           =   735
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Saldo Final:"
            Height          =   255
            Index           =   21
            Left            =   5400
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Importe:"
         Height          =   255
         Index           =   18
         Left            =   2760
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   5340
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Importe para Asignar:"
         Height          =   255
         Index           =   14
         Left            =   -73800
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   4980
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Grabar"
         Height          =   615
         Left            =   -71400
         Picture         =   "frmajusteproveedores.frx":05F5
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   5880
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmajusteproveedores.frx":0B27
         Height          =   255
         Left            =   -74760
         TabIndex        =   40
         Top             =   4980
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
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
      Begin VB.Frame Frame6 
         Caption         =   "Notas de Credito a Asignar"
         Height          =   1815
         Left            =   -74880
         TabIndex        =   35
         Top             =   3840
         Width           =   8175
         Begin VB.CommandButton Command5 
            Caption         =   "Total Comp.:"
            Height          =   255
            Index           =   9
            Left            =   4440
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Saldo Comp.:"
            Height          =   255
            Index           =   10
            Left            =   4440
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Importe a Asignar:"
            Height          =   255
            Index           =   11
            Left            =   4200
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   14
            Left            =   5760
            TabIndex        =   39
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   13
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   12
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   360
            Width           =   1455
         End
         Begin MSDataListLib.DataCombo DataCombo7 
            Bindings        =   "frmajusteproveedores.frx":0B43
            Height          =   315
            Left            =   1080
            TabIndex        =   36
            Top             =   360
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            ListField       =   "comp"
            BoundColumn     =   "comp"
            Text            =   ""
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmajusteproveedores.frx":0B5F
         Height          =   855
         Left            =   -70200
         TabIndex        =   31
         Top             =   4800
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   1508
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
      Begin VB.CommandButton Command2 
         Caption         =   "&Grabar"
         Height          =   615
         Left            =   -71880
         Picture         =   "frmajusteproveedores.frx":0B7B
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   5820
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   -71640
         TabIndex        =   29
         Top             =   4980
         Width           =   1095
      End
      Begin VB.Frame Frame5 
         Caption         =   "Cobro para Asignar"
         Height          =   855
         Left            =   -74880
         TabIndex        =   26
         Top             =   3840
         Width           =   8175
         Begin VB.CommandButton Command5 
            Caption         =   "Nº Orden:"
            Height          =   255
            Index           =   12
            Left            =   600
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Tot. Comp.:"
            Height          =   255
            Index           =   13
            Left            =   4800
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   11
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   360
            Width           =   1095
         End
         Begin MSDataListLib.DataCombo DataCombo5 
            Bindings        =   "frmajusteproveedores.frx":10AD
            Height          =   315
            Left            =   1920
            TabIndex        =   28
            Top             =   360
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            ListField       =   "nrorden"
            BoundColumn     =   "importe"
            Text            =   ""
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Grabar"
         Height          =   615
         Left            =   3720
         Picture         =   "frmajusteproveedores.frx":10C9
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   6000
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   5340
         Width           =   1215
      End
      Begin VB.ListBox List1 
         Height          =   1860
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   20
         Top             =   5160
         Width           =   2535
      End
      Begin VB.Frame Frame4 
         Caption         =   "Asignar a Proveedor"
         Height          =   1215
         Left            =   120
         TabIndex        =   18
         Top             =   3840
         Width           =   8175
         Begin VB.CommandButton Command5 
            Caption         =   "Comprobante:"
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Saldo:"
            Height          =   255
            Index           =   16
            Left            =   5760
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   840
            Width           =   735
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Tot.Comp.:"
            Height          =   255
            Index           =   17
            Left            =   3600
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   9
            Left            =   6480
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   810
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   8
            Left            =   4560
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   810
            Width           =   1095
         End
         Begin MSDataListLib.DataCombo DataCombo2 
            Bindings        =   "frmajusteproveedores.frx":15FB
            Height          =   315
            Left            =   720
            TabIndex        =   19
            Top             =   360
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            ListField       =   "razonsocial"
            BoundColumn     =   "codproveedor"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DataCombo3 
            Bindings        =   "frmajusteproveedores.frx":1615
            Height          =   315
            Left            =   1320
            TabIndex        =   22
            Top             =   780
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            ListField       =   "comp"
            BoundColumn     =   "comp"
            Text            =   ""
         End
      End
      Begin VB.CommandButton grabar 
         Caption         =   "&Grabar"
         Height          =   615
         Left            =   -71760
         Picture         =   "frmajusteproveedores.frx":162F
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4800
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Caption         =   "Asignar a Proveedor"
         Height          =   855
         Left            =   -74880
         TabIndex        =   1
         Top             =   3840
         Width           =   8175
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "frmajusteproveedores.frx":1B61
            Height          =   315
            Left            =   720
            TabIndex        =   2
            Top             =   360
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            ListField       =   "razonsocial"
            BoundColumn     =   "cuit"
            Text            =   ""
         End
      End
      Begin MSAdodcLib.Adodc datlibroventas 
         Height          =   330
         Left            =   -74880
         Top             =   -60
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
      Begin MSAdodcLib.Adodc datclientes 
         Height          =   330
         Left            =   -74880
         Top             =   300
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
      Begin MSAdodcLib.Adodc datrecibosabonan 
         Height          =   330
         Left            =   120
         Top             =   60
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
      Begin MSAdodcLib.Adodc datrecibos 
         Height          =   330
         Left            =   1440
         Top             =   60
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
      Begin MSAdodcLib.Adodc datasigcomp 
         Height          =   330
         Left            =   2760
         Top             =   60
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
      Begin MSAdodcLib.Adodc datrecibosinc 
         Height          =   330
         Left            =   -74880
         Top             =   60
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
      Begin MSAdodcLib.Adodc datrecibosabonan1 
         Height          =   330
         Left            =   -73560
         Top             =   60
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
      Begin MSAdodcLib.Adodc datlibroventas1 
         Height          =   330
         Left            =   -73560
         Top             =   -60
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
      Begin VB.CommandButton grabaasiento 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -71640
         TabIndex        =   70
         Top             =   5640
         Width           =   135
      End
      Begin MSAdodcLib.Adodc datasiento 
         Height          =   330
         Left            =   -74640
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
         LockType        =   2
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
         UserName        =   "lucva"
         Password        =   "25072004"
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
         Left            =   -73320
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
         LockType        =   2
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
         UserName        =   "lucva"
         Password        =   "25072004"
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
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "frmajusteproveedores.frx":1B7B
         Height          =   975
         Left            =   -74880
         TabIndex        =   71
         Top             =   5880
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1720
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
      Begin MSDataGridLib.DataGrid DataGrid4 
         Bindings        =   "frmajusteproveedores.frx":1B94
         Height          =   975
         Left            =   -70080
         TabIndex        =   72
         Top             =   5880
         Visible         =   0   'False
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   1720
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
      Begin MSAdodcLib.Adodc datasiento1 
         Height          =   330
         Left            =   -75000
         Top             =   600
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
      LcK1            =   "3.66*/4/0*/1-5*210/."
      LcK2            =   $"frmajusteproveedores.frx":1BAE
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
Attribute VB_Name = "frmajusteproveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim importeorden(200) As Currency
Dim idorden(200) As Double
Dim numorden As String
Dim max As Integer


Private Sub buscar_Click()
On Error Resume Next
List1.Clear

If SSTab1.Tab = 0 Or SSTab1.Tab = 2 Or SSTab1.Tab = 3 Or SSTab1.Tab = 4 Then
        datlibroventas.RecordSource = "select librocompras.* from librocompras WHERE empresa = " & login.empresaact & " and proveedor = '" & DataCombo4.Text & "' and tipocompr = '" & Combo1.Text & "' and numcompr = '" & MaskEdBox1.Text & "' "
        datlibroventas.Refresh
        If datlibroventas.Recordset.EOF = True Then
            mensa = MsgBox("Comprobante no existe", vbCritical, "Error")
            Unload Me
            frmajusteproveedores.Show
            Exit Sub
        End If
        
        Text1(0).Text = datlibroventas.Recordset.Fields("fecha")
        Text1(1).Text = datlibroventas.Recordset.Fields("proveedor")
        Text1(2).Text = datlibroventas.Recordset.Fields("total")
        Text1(2).Text = Format(Text1(2).Text, "##0.00")
        If IsNull(datlibroventas.Recordset.Fields("saldo")) = False Then
            Text1(3).Text = datlibroventas.Recordset.Fields("saldo")
        Else
            Text1(3).Text = Text1(2).Text
        End If
        Text1(3).Text = Format(Text1(3).Text, "##0.00")
        Text1(4).Text = datlibroventas.Recordset.Fields("asiento")
        Text1(5).Text = Str(datlibroventas.Recordset.Fields("inicioper")) + " -- " + Str(datlibroventas.Recordset.Fields("finper"))
        Text1(6).Text = datlibroventas.Recordset.Fields("cuit")
        
        If SSTab1.Tab = 2 Then
             datclientes.RecordSource = "select proveedores.* from proveedores where empresa = " & login.empresaact & " and razonsocial = '" & datlibroventas.Recordset.Fields("proveedor") & "' "
             datclientes.Refresh
            
             datrecibosinc.RecordSource = "select consultaordensinc.* from consultaordensinc where codproveedor = " & datclientes.Recordset.Fields("codproveedor") & ""
             datrecibosinc.Refresh
             
            datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where empresa = " & login.empresaact & " and perinicial = '" & datlibroventas.Recordset.Fields("inicioper") & "' and nroasiento = " & datlibroventas.Recordset.Fields("asiento") & " "
            datmaestro.Refresh
            If datmaestro.Recordset.EOF = True Then
                idmast = 0
            Else
                idmast = datmaestro.Recordset.Fields("idmasterasientos")
            End If
        
            datasiento1.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where empresa = " & login.empresaact & " and idmasterasientos = " & idmast & " and haber <> 0  "
            datasiento1.Refresh
            
        End If
        
        If SSTab1.Tab = 3 Then
             datclientes.RecordSource = "select proveedores.* from proveedores where empresa = " & login.empresaact & " and razonsocial = '" & datlibroventas.Recordset.Fields("proveedor") & "' "
             datclientes.Refresh
             
             datrecibosinc.RecordSource = "select conceptosabonan.* from conceptosabonan where empresa = " & login.empresaact & " and proveedor = '" & datclientes.Recordset.Fields("razonsocial") & "' and (tipocompr = 'NCA' or tipocompr = 'NCB' or tipocompr = 'NCC' or tipocompr = 'NDA' OR tipocompr = 'NDB' or tipocompr = 'NDC')  order by comp"
             datrecibosinc.Refresh

        End If
        
End If

If SSTab1.Tab = 1 Then
    If Text2.Text <> " " Then
        datlibroventas.RecordSource = "select librocompras.* from librocompras WHERE empresa = " & login.empresaact & " and tipocompr = '" & Combo1.Text & "' and numcompr = '" & MaskEdBox1.Text & "' and proveedor = '" & DataCombo4.Text & "'"
        datlibroventas.Refresh
    Else
        datlibroventas.RecordSource = "select librocompras.* from librocompras WHERE empresa = " & login.empresaact & " and tipocompr = '" & Text2.Text & "' and numcompr = '" & DataCombo6.Text & "' and proveedor = '" & DataCombo4.Text & "' "
        datlibroventas.Refresh
    End If
        
        If datlibroventas.Recordset.EOF = True Then
            mensa = MsgBox("Comprobante no existe", vbCritical, "Error")
            Unload Me
            frmajusteproveedores.Show
            Exit Sub
        End If
        
        Text1(0).Text = datlibroventas.Recordset.Fields("fecha")
        Text1(1).Text = datlibroventas.Recordset.Fields("proveedor")
        Text1(2).Text = datlibroventas.Recordset.Fields("total")
        Text1(2).Text = Format(Text1(2).Text, "##0.00")
        If IsNull(datlibroventas.Recordset.Fields("saldo")) = False Then
            Text1(3).Text = datlibroventas.Recordset.Fields("saldo")
        Else
            Text1(3).Text = Text1(2).Text
        End If
        Text1(3).Text = Format(Text1(3).Text, "##0.00")
        Text1(4).Text = datlibroventas.Recordset.Fields("asiento")
        Text1(5).Text = Str(datlibroventas.Recordset.Fields("inicioper")) + " -- " + Str(datlibroventas.Recordset.Fields("finper"))
        Text1(6).Text = datlibroventas.Recordset.Fields("cuit")
        
        compro = Combo1.Text + "  " + MaskEdBox1.Text
        datrecibosabonan.RecordSource = "select ordendepagoabonan.* from ordendepagoabonan WHERE empresa = " & login.empresaact & " and comprobante = '" & compro & "' and nomproveedor = '" & Text1(1).Text & "'  "
        datrecibosabonan.Refresh
        
        If datrecibosabonan.Recordset.EOF = True Then Exit Sub
        List1.Clear
        
        i = 0
        datrecibosabonan.Recordset.MoveFirst
        Do While Not datrecibosabonan.Recordset.EOF
            datrecibos.RecordSource = "select ordendepago.* from ordendepago where empresa = " & login.empresaact & " and nrorden = '" & datrecibosabonan.Recordset.Fields("nrorden") & "' and anulado <> 'S'"
            datrecibos.Refresh
            
            If datrecibos.Recordset.EOF = False Then
                List1.AddItem (datrecibos.Recordset.Fields("nrorden"))
                importeorden(i) = datrecibosabonan.Recordset.Fields("importe")
                idorden(i) = datrecibosabonan.Recordset.Fields("id")
            End If
            max = i
            i = i + 1
            datrecibosabonan.Recordset.MoveNext
        Loop
        
End If



End Sub

Private Sub buscar_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        If SSTab1.Tab = 0 Then DataCombo1.SetFocus
        If SSTab1.Tab = 1 Then DataCombo2.SetFocus
        If SSTab1.Tab = 2 Then DataCombo4.SetFocus
    End If

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        MaskEdBox1.SetFocus
    End If

End Sub

Private Sub Command1_Click()
On Error GoTo fuera
    
    If Text2.Text = " " Then GoTo ajustes
    
    If DataCombo3.Text = "" Then GoTo sincomp
   
    If Val(Text1(7).Text) > Val(Text1(9).Text) Then
        mensa = MsgBox("No se puede realizar esta operacion porque el Importe a Asignar es mayor que el SALDO del Comprobante", vbCritical, "Error")
        Exit Sub
    End If

sincomp:
    compro = Combo1.Text + "  " + MaskEdBox1.Text
    If DataCombo2.Text = "" Then
        mensa = MsgBox("Ingrese el Proveedor de Destino", vbCritical, "Error")
        DataCombo2.SetFocus
        Exit Sub
    End If
    datrecibosabonan.RecordSource = "select ordendepagoabonan.* from ordendepagoabonan WHERE empresa = " & login.empresaact & " and id = '" & numorden & "'  "
    datrecibosabonan.Refresh
    
    If datrecibosabonan.Recordset.EOF = True Then GoTo fuera
    
    datrecibosabonan.Recordset.Fields("nomproveedor") = DataCombo2.Text
    datrecibosabonan.Recordset.Fields("codproveedor") = DataCombo2.BoundText
    
    If DataCombo3.Text <> "" Then
        datrecibosabonan.Recordset.Fields("comprobante") = DataCombo3.Text
    Else
            datrecibosabonan.Recordset.Fields(7) = Null
            datrecibosabonan.Recordset.Fields(9) = Val(Text1(7).Text)
            datrecibosabonan.Recordset.UpdateBatch adAffectCurrent
    End If
    
    datrecibosabonan.Recordset.UpdateBatch adAffectCurrent
    datlibroventas.Recordset.Fields("saldo") = Val(Text1(7).Text) + Val(Text1(3).Text)
    datlibroventas.Recordset.UpdateBatch adAffectCurrent
    

    If DataCombo3.Text <> "" Then
        datlibroventas.RecordSource = "select librocompras.* from librocompras WHERE empresa = " & login.empresaact & " and proveedor ='" & DataCombo2.Text & "' and tipocompr = '" & Left(DataCombo3.Text, 3) & "' and numcompr = '" & Right(DataCombo3.Text, 13) & "' "
        datlibroventas.Refresh
        datlibroventas.Recordset.Fields("saldo") = Val(Text1(9).Text) - Val(Text1(7).Text)
        datlibroventas.Recordset.Fields("imputado") = "S"
        datlibroventas.Recordset.UpdateBatch adAffectCurrent
    End If
    GoTo auditoria
    
ajustes:
    If Val(Text1(7).Text) > Val(Text1(2).Text) Then
        mensa = MsgBox("No se puede realizar esta operacion porque el Importe a Asignar es mayor que el Importe del Comprobante", vbCritical, "Error")
        Exit Sub
    End If
    If DataCombo2.Text = "" Then
        mensa = MsgBox("Ingrese el proveedor de Destino", vbCritical, "Error")
        DataCombo2.SetFocus
        Exit Sub
    End If
    datlibroventas.Recordset.Fields("saldo") = Val(Text1(3).Text) - Val(Text1(7).Text)
    datlibroventas.Recordset.Fields("imputado") = "S"
    datlibroventas.Recordset.UpdateBatch adAffectCurrent
    
    
    datlibroventas1.RecordSource = "select librocompras.* from librocompras WHERE empresa = " & login.empresaact & " and proveedor ='" & DataCombo2.Text & "' and numcompr = '" & DataCombo6.Text & "' "
    datlibroventas1.Refresh

If datlibroventas1.Recordset.EOF = True Then
    datlibroventas1.Recordset.AddNew
    For x = 1 To 66
        datlibroventas1.Recordset.Fields(x) = datlibroventas.Recordset.Fields(x)
    Next x
        datlibroventas1.Recordset.Fields("total") = Text1(7).Text
        datlibroventas1.Recordset.Fields("saldo") = Null
        datlibroventas1.Recordset.Fields("imputado") = Null
        datlibroventas1.Recordset.Fields("proveedor") = DataCombo2.Text
        datlibroventas1.Recordset.UpdateBatch adAffectCurrent
Else
        datlibroventas1.Recordset.Fields("total") = datlibroventas1.Recordset.Fields("total") + Text1(7).Text
        If IsNull(datlibroventas1.Recordset.Fields("saldo")) = False Then
            datlibroventas1.Recordset.Fields("saldo") = datlibroventas1.Recordset.Fields("saldo") + Val(Text1(7).Text)
        End If
        datlibroventas1.Recordset.UpdateBatch adAffectCurrent
End If

auditoria:
                Inicio.datauditoria.ConnectionString = login.conexiontotal
    
                Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
                Inicio.datauditoria.Refresh
    
                Inicio.datauditoria.Recordset.AddNew
                Inicio.datauditoria.Recordset.Fields("fecha") = Date
                Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
                Inicio.datauditoria.Recordset.Fields("ventana") = "AJUSTES EC PROVEEDORES"
                Inicio.datauditoria.Recordset.Fields("accion") = "Cambio de Cobro en FACT:" + compro
                Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
                Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
                Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent
    
    mensa = MsgBox("Cambio realizado", vbInformation, "Accion")
    
    
    
    Call Command6_Click
    Exit Sub

fuera:
    mensa = MsgBox("No se realizo el cambio", vbCritical, "Error")
    

End Sub

Private Sub Command2_Click()
On Error GoTo fuera

    If Val(Text1(10).Text) > Val(Text1(11).Text) Then
        mensa = MsgBox("No se puede realizar esta operacion porque el Importe a Asignar es mayor que el Importe del Comprobante", vbCritical, "Error")
        Text1(10).SetFocus
        Exit Sub
    End If

    If Val(Text1(10).Text) > Val(Text1(3).Text) Then
        mensa = MsgBox("No se puede realizar esta operacion porque el Importe a Asignar es mayor que el Saldo de la Factura", vbCritical, "Error")
        Text1(10).SetFocus
        Exit Sub
    End If
    
    identi = DataGrid1.Columns(10).Text
    
        datrecibosabonan.RecordSource = "select ordendepagoabonan.* from ordendepagoabonan WHERE empresa = " & login.empresaact & " and id = " & identi & ""
        datrecibosabonan.Refresh
        datrecibosabonan1.RecordSource = "select ordendepagoabonan.* from ordendepagoabonan WHERE empresa = " & login.empresaact & ""
        datrecibosabonan1.Refresh
        compro = Combo1.Text + "  " + MaskEdBox1.Text
        
        datrecibosabonan.Recordset.Fields("comprobante") = compro
        datrecibosabonan.Recordset.Fields("importe") = Text1(10).Text
        datrecibosabonan.Recordset.UpdateBatch adAffectCurrent
        If Val(Text1(10).Text) < Val(Text1(11).Text) Then
            datrecibosabonan1.Recordset.AddNew
            datrecibosabonan1.Recordset.Fields(0) = datrecibosabonan.Recordset.Fields(0)
            datrecibosabonan1.Recordset.Fields(1) = datrecibosabonan.Recordset.Fields(1)
            datrecibosabonan1.Recordset.Fields(2) = datrecibosabonan.Recordset.Fields(2)
            datrecibosabonan1.Recordset.Fields(3) = datrecibosabonan.Recordset.Fields(3)
            For x = 5 To 9
                datrecibosabonan1.Recordset.Fields(x) = datrecibosabonan.Recordset.Fields(x)
            Next x
            datrecibosabonan1.Recordset.Fields(7) = Null
            datrecibosabonan1.Recordset.Fields(9) = Val(Text1(11).Text) - Val(Text1(10).Text)
            datrecibosabonan1.Recordset.UpdateBatch adAffectCurrent
        End If
        datlibroventas.Recordset.Fields("saldo") = Val(Text1(3).Text) - Val(Text1(10).Text)
        datlibroventas.Recordset.Fields("imputado") = "S"
        datlibroventas.Recordset.UpdateBatch adAffectCurrent
                    
                    
        cuentahaber = DataGrid3.Columns(2).Text
        cuentadebe = DataGrid4.Columns(2).Text
        importeasginado = Val(Text1(10).Text)
        If cuentadebe <> cuentahaber Then
                datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where empresa = " & login.empresaact & " and perinicial = '" & login.iper & "' order by nroasiento"
                datmaestro.Refresh
                If datmaestro.Recordset.EOF = True Then
                    ultimoasiento = 1
                Else
                    datmaestro.Recordset.MoveLast
                    ultimoasiento = datmaestro.Recordset.Fields(3) + 1
                End If
                datmaestro.Recordset.AddNew
                datmaestro.Recordset.Fields("fecha") = Text1(0).Text
                datmaestro.Recordset.Fields("fecharegistro") = Text1(0).Text
                datmaestro.Recordset.Fields("nroasiento") = ultimoasiento
                datmaestro.Recordset.Fields("concepto") = "Asignacion Orden:" + DataCombo5.Text
                datmaestro.Recordset.Fields("perinicial") = login.iper
                datmaestro.Recordset.Fields("perfinal") = login.fper
                datmaestro.Recordset.Fields("empresa") = login.empresaact
                datmaestro.Recordset.Fields("libro") = "A"
                datmaestro.Recordset.Fields(11) = "S"
        
                datmaestro.Recordset.UpdateBatch adAffectCurrent
                
                datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where empresa = '0'"
                datasiento.Refresh
            
                datasiento.Recordset.AddNew
                datasiento.Recordset.Fields("idmasterasientos") = datmaestro.Recordset.Fields("idmasterasientos")
                datasiento.Recordset.Fields("fecha") = Text1(0).Text
                datasiento.Recordset.Fields("empresa") = login.empresaact
                datasiento.Recordset.Fields("idcuenta") = cuentadebe
                datasiento.Recordset.Fields("debe") = importeasginado
                datasiento.Recordset.Fields("haber") = 0
                datasiento.Recordset.Fields("detallefila") = "Orden:" + DataCombo5.Text
                datasiento.Recordset.UpdateBatch adAffectCurrent
                
                datasiento.Recordset.AddNew
                datasiento.Recordset.Fields("idmasterasientos") = datmaestro.Recordset.Fields("idmasterasientos")
                datasiento.Recordset.Fields("fecha") = Text1(0).Text
                datasiento.Recordset.Fields("empresa") = login.empresaact
                datasiento.Recordset.Fields("idcuenta") = cuentahaber
                datasiento.Recordset.Fields("debe") = 0
                datasiento.Recordset.Fields("haber") = importeasginado
                datasiento.Recordset.Fields("detallefila") = "Orden:" + DataCombo5.Text
                datasiento.Recordset.UpdateBatch adAffectCurrent
               
                                              
        End If


 
        Inicio.datauditoria.ConnectionString = login.conexiontotal
    
                Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
                Inicio.datauditoria.Refresh
    
                Inicio.datauditoria.Recordset.AddNew
                Inicio.datauditoria.Recordset.Fields("fecha") = Date
                Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
                Inicio.datauditoria.Recordset.Fields("ventana") = "AJUSTES EC PROVEEDORES"
                Inicio.datauditoria.Recordset.Fields("accion") = "Asig.de Cobro en FACT:" + compro
                Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
                Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
                Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent
    
    mensa = MsgBox("Cambio realizado", vbInformation, "Accion")
    
    
    
    Call Command6_Click
    Exit Sub

fuera:
    mensa = MsgBox("No se realizo el cambio", vbCritical, "Error")
    


End Sub

Private Sub Command3_Click()
On Error GoTo fuera

    If Val(Text1(14).Text) > Val(Text1(13).Text) Then
        mensa = MsgBox("No se puede realizar esta operacion porque el Importe a Asignar es mayor que el Saldo del Comprobante", vbCritical, "Error")
        Text1(10).SetFocus
        Exit Sub
    End If

    If Val(Text1(14).Text) > Val(Text1(3).Text) Then
        mensa = MsgBox("No se puede realizar esta operacion porque el Importe a Asignar es mayor que el Saldo de la Factura", vbCritical, "Error")
        Text1(14).SetFocus
        Exit Sub
    End If
    
   
        datrecibosabonan.RecordSource = "select ordendepagoabonan.* from ordendepagoabonan WHERE empresa = " & login.empresaact & " "
        datrecibosabonan.Refresh
        
        compro = Combo1.Text + "  " + MaskEdBox1.Text
        
        datrecibosabonan.Recordset.AddNew
        datrecibosabonan.Recordset.Fields("nrorden") = DataCombo7.Text
        datrecibosabonan.Recordset.Fields("empresa") = login.empresaact
        datrecibosabonan.Recordset.Fields("inicioper") = DataGrid2.Columns(2)
        datrecibosabonan.Recordset.Fields("finper") = DataGrid2.Columns(3)
        datrecibosabonan.Recordset.Fields("codproveedor") = datclientes.Recordset.Fields("codproveedor")
        datrecibosabonan.Recordset.Fields("nomproveedor") = datclientes.Recordset.Fields("razonsocial")
        datrecibosabonan.Recordset.Fields("codproveedor") = datclientes.Recordset.Fields("codproveedor")
        datrecibosabonan.Recordset.Fields("comprobante") = compro
        datrecibosabonan.Recordset.Fields("fechacompro") = DataGrid2.Columns(0)
        datrecibosabonan.Recordset.Fields("importe") = Text1(14).Text
        datrecibosabonan.Recordset.Fields("saldofactura") = 0
        datrecibosabonan.Recordset.UpdateBatch adAffectCurrent
        
        datlibroventas.Recordset.Fields("saldo") = Val(Text1(3).Text) - Val(Text1(14).Text)
        datlibroventas.Recordset.Fields("imputado") = "S"
        datlibroventas.Recordset.UpdateBatch adAffectCurrent
                    
        datlibroventas.RecordSource = "select librocompras.* from librocompras WHERE empresa = " & login.empresaact & " and tipocompr = '" & Left(DataCombo7.Text, 3) & "' and numcompr = '" & Right(DataCombo7.Text, 13) & "' and proveedor = '" & datclientes.Recordset.Fields("razonsocial") & "' "
        datlibroventas.Refresh
        datlibroventas.Recordset.Fields("saldo") = Val(Text1(14).Text) - Val(Text1(13).Text)
        datlibroventas.Recordset.Fields("imputado") = "S"
        datlibroventas.Recordset.UpdateBatch adAffectCurrent
 
        Inicio.datauditoria.ConnectionString = login.conexiontotal
    
                Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
                Inicio.datauditoria.Refresh
    
                Inicio.datauditoria.Recordset.AddNew
                Inicio.datauditoria.Recordset.Fields("fecha") = Date
                Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
                Inicio.datauditoria.Recordset.Fields("ventana") = "AJUSTES EC PROVEEDORES"
                Inicio.datauditoria.Recordset.Fields("accion") = "Asig.de Nota Credito a FACT:" + compro
                Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
                Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
                Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent
    
    mensa = MsgBox("Cambio realizado", vbInformation, "Accion")
    
    
    
    Call Command6_Click
    Exit Sub

fuera:
    mensa = MsgBox("No se realizo el cambio", vbCritical, "Error")


End Sub

Private Sub Command4_Click()
On Error Resume Next
   
    mensa = MsgBox("Graba Asiento Contable", vbYesNo, "Asiento")
   
    datrecibosabonan.RecordSource = "select ordendepagoabonan.* from ordendepagoabonan WHERE empresa = " & login.empresaact & " "
    datrecibosabonan.Refresh
       
    datrecibosabonan.Recordset.AddNew
    datrecibosabonan.Recordset.Fields("nrorden") = "Aj-" + Combo1.Text + MaskEdBox1.Text
    datrecibosabonan.Recordset.Fields("empresa") = login.empresaact
    datrecibosabonan.Recordset.Fields("inicioper") = login.iper
    datrecibosabonan.Recordset.Fields("finper") = login.fper
    datrecibosabonan.Recordset.Fields("codproveedor") = DataCombo4.BoundText
    datrecibosabonan.Recordset.Fields("nomproveedor") = DataCombo4.Text
    datrecibosabonan.Recordset.Fields("comprobante") = Combo1.Text + "  " + MaskEdBox1.Text
    datrecibosabonan.Recordset.Fields("fechacompro") = Date
    If Text3.Text = "" Then Text3.Text = 0
    If Text6.Text = "" Then Text6.Text = 0
    If Text3.Text <> 0 Then
        datrecibosabonan.Recordset.Fields("importe") = Text3.Text * -1
    Else
        datrecibosabonan.Recordset.Fields("importe") = Text6.Text
    End If
    datrecibosabonan.Recordset.Fields("codcuenta") = datlibroventas.Recordset.Fields("cht")
    datrecibosabonan.Recordset.Fields("saldofactura") = 0
    datrecibosabonan.Recordset.UpdateBatch adAffectCurrent
    
    datlibroventas.Recordset.Fields("saldo") = Text4.Text
    datlibroventas.Recordset.Fields("imputado") = "S"
    datlibroventas.Recordset.UpdateBatch adAffectCurrent
    
    If mensa = vbYes Then
        z_cuentas.menucuentas = "busca1"
        z_cuentas.Show
        Exit Sub
    End If
    
        mensa = MsgBox("Ajuste realizado", vbInformation, "Accion")
        
    
    Call Command6_Click
    Exit Sub
End Sub

Private Sub Command6_Click()

       ubicac = SSTab1.Tab
    
       Unload Me
       frmajusteproveedores.Show
       SSTab1.Tab = ubicac

End Sub

Private Sub DataCombo2_Click(Area As Integer)

    datasigcomp.RecordSource = "select conceptosabonan.* from conceptosabonan where empresa = " & login.empresaact & " and proveedor = '" & DataCombo2.Text & "'  order by comp"
    datasigcomp.Refresh
    DataCombo3.Text = ""

End Sub

Private Sub DataCombo2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        DataCombo3.SetFocus
    End If
    

End Sub

Private Sub DataCombo2_KeyUp(KeyCode As Integer, Shift As Integer)

    datasigcomp.RecordSource = "select conceptosabonan.* from conceptosabonan where empresa = " & login.empresaact & " and proveedor = '" & DataCombo2.Text & "' order by comp "
    datasigcomp.Refresh
    DataCombo3.Text = ""

End Sub

Private Sub DataCombo3_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then
        KeyAscii = 0
        datasigcomp.RecordSource = "select conceptosabonan.* from conceptosabonan where empresa = " & login.empresaact & " and proveedor = '" & DataCombo2.Text & "' and comp = '" & DataCombo3.Text & "' order by comp "
        datasigcomp.Refresh
    
        Text1(8).Text = datasigcomp.Recordset.Fields("total")
        Text1(8).Text = Format(Text1(8).Text, "##0.00")
        If IsNull(datasigcomp.Recordset.Fields("saldo")) = False Then
            Text1(9).Text = datasigcomp.Recordset.Fields("saldo")
        Else
            Text1(9).Text = Text1(8).Text
        End If
        Text1(9).Text = Format(Text1(9).Text, "##0.00")
        datasigcomp.RecordSource = "select conceptosabonan.* from conceptosabonan where empresa = " & login.empresaact & " and proveedor = '" & DataCombo2.Text & "' order by comp"
        datasigcomp.Refresh
    End If

End Sub

Private Sub DataCombo4_Click(Area As Integer)

        datlibroventas.RecordSource = "select librocompras.* from librocompras WHERE empresa = " & login.empresaact & " and proveedor = '" & DataCombo4.Text & "'"
        datlibroventas.Refresh


End Sub

Private Sub DataCombo4_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    KeyAscii = 0
    DataCombo6.SetFocus
End If


End Sub

Private Sub DataCombo4_KeyUp(KeyCode As Integer, Shift As Integer)

        datlibroventas.RecordSource = "select librocompras.* from librocompras WHERE empresa = " & login.empresaact & " and proveedor = '" & DataCombo4.Text & "'"
        datlibroventas.Refresh


End Sub

Private Sub DataCombo5_Click(Area As Integer)
On Error Resume Next

    Text1(11).Text = DataCombo5.BoundText
    Text1(11).Text = Format(Text1(11).Text, "##0.00")
    Text1(10).Text = Text1(11).Text
    If IsNull(DataCombo5.SelectedItem) = False Then
        DataGrid1.Bookmark = DataCombo5.SelectedItem
        datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where empresa = " & login.empresaact & " and perinicial = '" & DataGrid1.Columns(8).Text & "' and nroasiento = " & DataGrid1.Columns(11) & " "
        datmaestro.Refresh
        If datmaestro.Recordset.EOF = True Then
            idmast = 0
        Else
            idmast = datmaestro.Recordset.Fields("idmasterasientos")
        End If
        
        datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where empresa = " & login.empresaact & " and idmasterasientos = " & idmast & " and debe <> 0  "
        datasiento.Refresh
    End If
    
    
End Sub

Private Sub DataCombo5_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(10).SetFocus
    End If

End Sub

Private Sub DataCombo5_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Text1(11).Text = DataCombo5.BoundText
    Text1(11).Text = Format(Text1(11).Text, "##0.00")
    Text1(10).Text = Text1(11).Text
    DataGrid1.Bookmark = DataCombo5.SelectedItem

End Sub

Private Sub DataCombo6_Click(Area As Integer)

    Text2.Text = DataCombo6.BoundText
    If Text2.Text = " " Then
        Text1(7).Locked = False
    Else
        Text1(7).Locked = True
    End If

End Sub

Private Sub DataCombo6_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then
        KeyAscii = 0

        Combo1.Text = Text2.Text
        If Text2.Text = " " Then
            MaskEdBox1.Mask = ""
            MaskEdBox1.Text = DataCombo6.Text
        Else
            MaskEdBox1.Mask = "####-########"
            MaskEdBox1.Text = DataCombo6.Text
        End If
        Call buscar_Click
        If SSTab1.Tab = 0 Then DataCombo1.SetFocus
        If SSTab1.Tab = 1 Then DataCombo2.SetFocus
        If SSTab1.Tab = 2 Then DataCombo5.SetFocus
        If SSTab1.Tab = 3 Then DataCombo7.SetFocus
        If SSTab1.Tab = 4 Then Text3.SetFocus
    End If
    

End Sub

Private Sub DataCombo6_KeyUp(KeyCode As Integer, Shift As Integer)
     
     Text2.Text = DataCombo6.BoundText
    If Text2.Text = " " Then
        Text1(7).Locked = False
    Else
        Text1(7).Locked = True
    End If
     
End Sub

Private Sub DataCombo7_Click(Area As Integer)
On Error Resume Next

    DataGrid2.Bookmark = DataCombo7.SelectedItem
    Text1(12).Text = DataGrid2.Columns(5).Text * -1
    If DataGrid2.Columns(8).Text <> "" Then
        Text1(13).Text = DataGrid2.Columns(8).Text * -1
    Else
        Text1(13).Text = Text1(12).Text
    End If
    Text1(14).Text = Text1(13).Text
    Text1(12).Text = Format(Text1(12).Text, "##0.00")
    Text1(13).Text = Format(Text1(13).Text, "##0.00")
    Text1(14).Text = Format(Text1(14).Text, "##0.00")


End Sub

Private Sub DataCombo7_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(14).SetFocus
    End If
    

End Sub

Private Sub DataCombo7_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    DataGrid2.Bookmark = DataCombo7.SelectedItem
    Text1(12).Text = DataGrid2.Columns(5).Text * -1
    If DataGrid2.Columns(8).Text <> "" Then
        Text1(13).Text = DataGrid2.Columns(8).Text * -1
    Else
        Text1(13).Text = Text1(12).Text
    End If
    Text1(14).Text = Text1(13).Text
    Text1(12).Text = Format(Text1(12).Text, "##0.00")
    Text1(13).Text = Format(Text1(13).Text, "##0.00")
    Text1(14).Text = Format(Text1(14).Text, "##0.00")


End Sub

Private Sub Form_Activate()
    If login.ajuestesec <> "S" Then
        MsgBox "No tiene permiso de Acceso a esta Opcion", vbInformation, "Denegado"
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
frmajusteproveedores.Top = 0
frmajusteproveedores.Left = 0

Aplicar_skin Me

    datlibroventas.ConnectionString = login.conexiontotal
    datlibroventas1.ConnectionString = login.conexiontotal
    datclientes.ConnectionString = login.conexiontotal
    datrecibosabonan.ConnectionString = login.conexiontotal
    datrecibos.ConnectionString = login.conexiontotal
    datasigcomp.ConnectionString = login.conexiontotal
    datrecibosinc.ConnectionString = login.conexiontotal
    datrecibosabonan1.ConnectionString = login.conexiontotal
    datasiento.ConnectionString = login.conexiontotal
    datasiento1.ConnectionString = login.conexiontotal
    datmaestro.ConnectionString = login.conexiontotal
    
    datasigcomp.RecordSource = "select conceptosabonan.* from conceptosabonan where empresa = " & login.empresaact & " order by comp"
    datasigcomp.Refresh
    
    datclientes.RecordSource = "select proveedores.* from proveedores where empresa = " & login.empresaact & " order by razonsocial"
    datclientes.Refresh
    
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
                    Combo1.AddItem "F-C"
                    Combo1.AddItem "R-C"
                    Combo1.AddItem "NDC"
                    Combo1.AddItem "NCC"
                    Combo1.AddItem "REC"
                    Combo1.AddItem "TKT"
                    Combo1.ListIndex = 0


End Sub

Private Sub grabaasiento_Click()

On Error GoTo fuera

    datmaestro.RecordSource = "select [Maestro Asientos].* from [Maestro Asientos] where empresa = " & login.empresaact & " and perinicial = '" & login.iper & "'"
    datmaestro.Refresh
    datasiento.RecordSource = "select [Detalle Asientos].* from [Detalle Asientos] where empresa = " & login.empresaact & ""
    datasiento.Refresh
    
    If datmaestro.Recordset.EOF = False Then

            datmaestro.Recordset.MoveLast
            nroasie = datmaestro.Recordset.Fields(3) + 1
    Else
            nroasie = 1
    End If
         
         
    If IsNull(datlibroventas.Recordset.Fields("cht")) = True Then GoTo fuera
    If Text5.Text = "" Then GoTo fuera
    
    
         
    datmaestro.Recordset.AddNew
    datmaestro.Recordset.Fields(0) = Date
    datmaestro.Recordset.Fields(1) = Date
    datmaestro.Recordset.Fields(3) = nroasie
    datmaestro.Recordset.Fields(4) = "Ajuste " + Combo1.Text + MaskEdBox1.Text
    datmaestro.Recordset.Fields(5) = login.iper
    datmaestro.Recordset.Fields(6) = login.fper
    datmaestro.Recordset.Fields(7) = login.empresaact
    datmaestro.Recordset.Fields(8) = "N"
    datmaestro.Recordset.Fields(10) = ""
    datmaestro.Recordset.Fields(11) = "S"
    datmaestro.Recordset.UpdateBatch adAffectCurrent

            datasiento.Recordset.AddNew
            datasiento.Recordset.Fields(1) = datmaestro.Recordset.Fields(0).Value
            datasiento.Recordset.Fields(2) = datlibroventas.Recordset.Fields("cht")
        If Val(Text3.Text) <> 0 Then
            If Val(Text3.Text) < 0 Then
                datasiento.Recordset.Fields(3) = Val(Text3.Text) * -1
                datasiento.Recordset.Fields(4) = 0
            Else
                datasiento.Recordset.Fields(3) = 0
                datasiento.Recordset.Fields(4) = Val(Text3.Text)
            End If
        Else
            If Val(Text6.Text) < 0 Then
                datasiento.Recordset.Fields(3) = 0
                datasiento.Recordset.Fields(4) = Val(Text6.Text) * -1
            Else
                datasiento.Recordset.Fields(3) = Val(Text6.Text)
                datasiento.Recordset.Fields(4) = 0
            End If
        End If
            datasiento.Recordset.Fields(5) = Val(datmaestro.Recordset.Fields(2))
            datasiento.Recordset.Fields(6) = "Ajuste " + Combo1.Text + MaskEdBox1.Text
            datasiento.Recordset.Fields(7) = login.empresaact
            datasiento.Recordset.Fields(8) = datlibroventas.Recordset.Fields("ccosto")
            datasiento.Recordset.UpdateBatch adAffectCurrent
            
            datasiento.Recordset.AddNew
            datasiento.Recordset.Fields(1) = datmaestro.Recordset.Fields(0).Value
            datasiento.Recordset.Fields(2) = Text5.Text
        If Val(Text3.Text) <> 0 Then
            If Val(Text3.Text) < 0 Then
                datasiento.Recordset.Fields(3) = 0
                datasiento.Recordset.Fields(4) = Val(Text3.Text) * -1
            Else
                datasiento.Recordset.Fields(3) = Val(Text3.Text)
                datasiento.Recordset.Fields(4) = 0
            End If
        Else
            If Val(Text6.Text) < 0 Then
                datasiento.Recordset.Fields(3) = Val(Text6.Text) * -1
                datasiento.Recordset.Fields(4) = 0
            Else
                datasiento.Recordset.Fields(3) = 0
                datasiento.Recordset.Fields(4) = Val(Text6.Text)
            End If
        End If
            datasiento.Recordset.Fields(5) = Val(datmaestro.Recordset.Fields(2))
            datasiento.Recordset.Fields(6) = "Ajuste " + Combo1.Text + MaskEdBox1.Text
            datasiento.Recordset.Fields(7) = login.empresaact
            datasiento.Recordset.Fields(8) = datlibroventas.Recordset.Fields("ccosto")
            datasiento.Recordset.UpdateBatch adAffectCurrent
            
        
    mensa = MsgBox("Ajuste realizado", vbInformation, "Accion")
          
    Unload Me
    frmajusteproveedores.Show
    Exit Sub

fuera:

    MsgBox "No se pudo registras el Asiento contable", vbCritical, "Verificar Minuta contable"
End Sub

Private Sub grabar_Click()
On Error GoTo fuera

    datlibroventas.Recordset.Fields("proveedor") = DataCombo1.Text
    datlibroventas.Recordset.Fields("cuit") = DataCombo1.BoundText
    datlibroventas.Recordset.UpdateBatch adAffectCurrent
    
    
                Inicio.datauditoria.ConnectionString = login.conexiontotal
    
                Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
                Inicio.datauditoria.Refresh
    
                Inicio.datauditoria.Recordset.AddNew
                Inicio.datauditoria.Recordset.Fields("fecha") = Date
                Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
                Inicio.datauditoria.Recordset.Fields("ventana") = "AJUSTES EC PROVEEDORES"
                Inicio.datauditoria.Recordset.Fields("accion") = "Cambio de proveedor en FACT:" + Combo1.Text + "-" + MaskEdBox1.Text
                Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
                Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
                Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent
    
    mensa = MsgBox("Cambio realizado", vbInformation, "Accion")
    
    
    
    Call Command6_Click
    Exit Sub

fuera:
    mensa = MsgBox("No se realizo el cambio", vbCritical, "Error")

End Sub

Private Sub List1_ItemCheck(Item As Integer)

If List1.ListIndex > 0 Then
    For x = 0 To List1.ListIndex - 1
        List1.Selected(x) = False
    Next x
    For x = List1.ListIndex + 1 To max
        List1.Selected(x) = False
    Next x
Else
    For x = List1.ListIndex + 1 To max
        List1.Selected(x) = False
    Next x
End If


    If List1.Selected(List1.ListIndex) = True Then
        Text1(7).Text = importeorden(List1.ListIndex)
        numorden = idorden(List1.ListIndex)
    Else
        Text1(7).Text = 0
    End If
    Text1(7).Text = Format(Text1(7).Text, "##0.00")

End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
            car = 0
            car1 = 0
            For x = 6 To 13
                If Mid(MaskEdBox1.Text, x, 1) = "_" Then
                    car = car + 1
                Else
                    car1 = car1 + 1
                End If
            Next x
            MaskEdBox1.Text = Mid(MaskEdBox1.Text, 1, 4) + "-" + Mid("0000000", 1, car) + Mid(MaskEdBox1.Text, 6, car1)
            buscar.SetFocus
    End If

End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)

    If DataCombo4.Text <> "" Then Call Command6_Click

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)


If KeyAscii <> 13 And KeyAscii <> 8 Then
    If KeyAscii < 46 Or KeyAscii > 57 Then KeyAscii = 0
End If

    If KeyAscii = 13 Then
        KeyAscii = 0
        If Index = 7 Then
            Text1(7).Text = Format(Text1(7).Text, "##0.00")
            Command1.SetFocus
        End If
        
        If Index = 10 Then
            Text1(10).Text = Format(Text1(10).Text, "##0.00")
            Command2.SetFocus
        End If
        If Index = 14 Then
            Text1(14).Text = Format(Text1(14).Text, "##0.00")
            Command3.SetFocus
        End If
    End If


End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
On Error Resume Next
       If KeyAscii = 13 Then
            KeyAscii = 0
            If Text3.Text < 0 Then Text3.Text = Text3.Text * -1
            Text4.Text = Val(Text1(3).Text) + Val(Text3.Text)
            Text3.Text = Format(Text3.Text, "##0.00")
             Command4.SetFocus
        End If

End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)

       If KeyAscii = 13 Then
            KeyAscii = 0
            If Text6.Text < 0 Then Text6.Text = Text6.Text * -1
            Text4.Text = Val(Text1(3).Text) - Val(Text6.Text)
            Text6.Text = Format(Text6.Text, "##0.00")
            Command4.SetFocus
        End If

End Sub
