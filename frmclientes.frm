VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form frmclientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clientes Maestro"
   ClientHeight    =   7515
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   10275
   Icon            =   "frmclientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7515
   ScaleWidth      =   10275
   Begin vbskpro.Skinner Skinner1 
      Left            =   120
      Top             =   6840
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
      Enabled         =   0   'False
      MinToBarButtonToolTipText=   "Minimizar a la barra de títulos"
      RestoreFromBarButtonToolTipText=   "Restaurar ventana"
      AlwaysOnTopButtonToolTipText=   "Hacer siempre visible"
      AlwaysOnTopDownButtonToolTipText=   "Hacer no siempre visible"
      ChangeSkinButtonToolTipText=   "Cambiar skin"
      HelpButtonToolTipText=   "Ayuda"
      SysEnableSkinCaption=   "Habilitar &Skin"
      SysDisableSkinCaption=   "Deshabilitar &Skin"
      LcK2            =   $"frmclientes.frx":11618
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
   Begin VB.CommandButton verificacuenta 
      Caption         =   "verificacuenta"
      Height          =   255
      Left            =   4080
      TabIndex        =   67
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton llena 
      Caption         =   "llena"
      Height          =   255
      Left            =   4440
      TabIndex        =   45
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmclientes.frx":11627
      Height          =   495
      Left            =   8640
      TabIndex        =   56
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
   Begin VB.Frame Frame3 
      Caption         =   "Impuestos"
      Height          =   1935
      Left            =   5400
      TabIndex        =   49
      Top             =   3000
      Width           =   4575
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   15
         Left            =   1920
         MaxLength       =   15
         TabIndex        =   17
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   14
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   16
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Nro.Inscripción 2:"
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
         Index           =   21
         Left            =   120
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Categoría 2:"
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
         Index           =   17
         Left            =   600
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Categoría 1:"
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
         Index           =   16
         Left            =   600
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Nro.Inscripción 1:"
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
         Index           =   15
         Left            =   120
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   13
         Left            =   1920
         MaxLength       =   15
         TabIndex        =   15
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   12
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   14
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Contabilidad"
      Height          =   1215
      Left            =   120
      TabIndex        =   46
      Top             =   3000
      Width           =   5175
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   11
         Left            =   2280
         TabIndex        =   13
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   10
         Left            =   2280
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cuenta de Ventas:"
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
         Index           =   14
         Left            =   240
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cuenta Cliente:"
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
         Index           =   13
         Left            =   240
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   360
         Width           =   1935
      End
      Begin MSAdodcLib.Adodc datcuentas 
         Height          =   330
         Left            =   3840
         Top             =   840
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
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   31
      Top             =   120
      Width           =   9855
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   7560
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   8
         Left            =   7560
         MaxLength       =   3
         TabIndex        =   11
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   7560
         MaxLength       =   13
         TabIndex        =   10
         Top             =   2040
         Width           =   1575
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmclientes.frx":11641
         Height          =   315
         Left            =   7560
         TabIndex        =   8
         Top             =   960
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "descripcion"
         BoundColumn     =   "categ"
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Có&digo:"
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
         Index           =   12
         Left            =   360
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   7
         Top             =   2400
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2040
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1680
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   3
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   2
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   4440
         MaxLength       =   50
         TabIndex        =   1
         Top             =   240
         Width           =   3975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Zona Fiscal:"
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
         Index           =   11
         Left            =   6240
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Numero:"
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
         Index           =   10
         Left            =   6600
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Tipo:"
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
         Index           =   9
         Left            =   6840
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Identificaion Impositiva"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   6240
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1320
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Categoria IVA:"
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
         Left            =   6000
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Contacto:"
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
         Left            =   360
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "e-mail:"
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
         Left            =   600
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Teléfono:"
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
         Left            =   360
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cod.Postal:"
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
         Left            =   240
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Localidad:"
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
         Left            =   360
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Domicilio:"
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
         Left            =   480
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Razon Social:"
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
         Left            =   2880
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSAdodcLib.Adodc datcondtrib 
      Height          =   330
      Left            =   0
      Top             =   3480
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
   Begin MSComctlLib.ImageList iml16 
      Left            =   0
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclientes.frx":1165B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmclientes.frx":1206D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Caption         =   "Otros"
      Height          =   3135
      Left            =   120
      TabIndex        =   52
      Top             =   4200
      Width           =   9855
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   24
         Left            =   2280
         MaxLength       =   20
         TabIndex        =   24
         Top             =   2400
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cobrador:"
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
         Index           =   28
         Left            =   1080
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   960
         Width           =   1095
      End
      Begin VB.Frame Frame5 
         Height          =   1095
         Left            =   4800
         TabIndex        =   65
         Top             =   2040
         Width           =   5055
         Begin KewlButtonz.KewlButtons elminar 
            Height          =   615
            Left            =   2040
            TabIndex        =   30
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1085
            BTYPE           =   14
            TX              =   "&Borrar"
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
            MICON           =   "frmclientes.frx":12A7F
            PICN            =   "frmclientes.frx":12A9B
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
            Left            =   1080
            TabIndex        =   29
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
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
            MICON           =   "frmclientes.frx":15E8D
            PICN            =   "frmclientes.frx":15EA9
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin KewlButtonz.KewlButtons aceptar 
            Height          =   615
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   1085
            BTYPE           =   14
            TX              =   "&Aceptar"
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
            MICON           =   "frmclientes.frx":168BB
            PICN            =   "frmclientes.frx":168D7
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin KewlButtonz.KewlButtons KewlButtons1 
            Height          =   615
            Left            =   3120
            TabIndex        =   68
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   1085
            BTYPE           =   14
            TX              =   "&Imprime"
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
            MICON           =   "frmclientes.frx":172E9
            PICN            =   "frmclientes.frx":17305
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin KewlButtonz.KewlButtons cerrar 
            Height          =   615
            Left            =   4080
            TabIndex        =   69
            Top             =   240
            Width           =   855
            _ExtentX        =   1508
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
            MICON           =   "frmclientes.frx":17D17
            PICN            =   "frmclientes.frx":17D33
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
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   23
         Left            =   7200
         MaxLength       =   3
         TabIndex        =   27
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   22
         Left            =   7200
         MaxLength       =   50
         TabIndex        =   26
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   21
         Left            =   7200
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   25
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Tipo Bonificación:"
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
         Index           =   27
         Left            =   5280
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Credito Máximo:"
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
         Index           =   26
         Left            =   5400
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Tipo de Cliente:"
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
         Index           =   25
         Left            =   5280
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   20
         Left            =   2280
         MaxLength       =   3
         TabIndex        =   22
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Motivo"
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
         Index           =   24
         Left            =   1200
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   2400
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Activo:"
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
         Index           =   23
         Left            =   1200
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   19
         Left            =   2280
         MaxLength       =   3
         TabIndex        =   21
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Transportista:"
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
         Index           =   22
         Left            =   600
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   18
         Left            =   2280
         MaxLength       =   3
         TabIndex        =   20
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   16
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   17
         Left            =   2280
         MaxLength       =   3
         TabIndex        =   19
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Vendedor:"
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
         Index           =   20
         Left            =   1080
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Condición de Venta:"
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
         Index           =   19
         Left            =   240
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Lista de Precios:"
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
         Index           =   18
         Left            =   600
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1575
      End
      Begin KewlButtonz.KewlButtons siguiente 
         Height          =   375
         Left            =   9000
         TabIndex        =   70
         TabStop         =   0   'False
         ToolTipText     =   "Siguiente"
         Top             =   1320
         Width           =   735
         _ExtentX        =   1296
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
         BCOL            =   -2147483629
         BCOLO           =   -2147483629
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmclientes.frx":1887D
         PICN            =   "frmclientes.frx":18899
         PICH            =   "frmclientes.frx":1F0FB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons anterior 
         Height          =   375
         Left            =   9000
         TabIndex        =   71
         TabStop         =   0   'False
         ToolTipText     =   "Anterior"
         Top             =   1680
         Width           =   735
         _ExtentX        =   1296
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
         BCOL            =   -2147483629
         BCOLO           =   -2147483629
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmclientes.frx":2595D
         PICN            =   "frmclientes.frx":25979
         PICH            =   "frmclientes.frx":2C1DB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin Crystal.CrystalReport CrystalReporte 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowTitle     =   "Libro IVA Compras"
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
      Connect         =   "PROVIDER=MSDASQL;dsn=contable;uid=sa;pwd=;database=contablesql;"
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
   Begin MSAdodcLib.Adodc datverifica 
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
End
Attribute VB_Name = "frmclientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim navega As Integer
Dim salir As Integer
Dim posicion As Integer
Dim empresareal As Integer
Public flagnuevo As Integer

Private Sub aceptar_Click()
On Error Resume Next
Dim importe As Currency
       
        If Text1(0).Text = "" Then
            MsgBox "No puede ingresar un valor Nulo en la RAZONSOCIAL", vbCritical, "Error"
            Text1(0).SetFocus
            Exit Sub
        End If
        
        If flagnuevo = 0 Then
                bases.datbasemenu.Recordset.AddNew
                bases.datbasemenu.Recordset.Fields("empresa") = login.empresaact
        End If
        
        bases.datbasemenu.Recordset.Fields("razonsocial") = Text1(0).Text
        bases.datbasemenu.Recordset.Fields("domicilio") = Text1(1).Text
        bases.datbasemenu.Recordset.Fields("localidad") = Text1(2).Text
        bases.datbasemenu.Recordset.Fields("codpostal") = Text1(3).Text
        bases.datbasemenu.Recordset.Fields("telefono") = Text1(4).Text
        bases.datbasemenu.Recordset.Fields("email") = Text1(5).Text
        bases.datbasemenu.Recordset.Fields("contacto") = Text1(6).Text
        If DataCombo1.Text = "" Then
            MsgBox "Ingrese Condición Tributaria", vbCritical, "Error"
            DataCombo1.SetFocus
            bases.datbasemenu.Recordset.Delete adAffectCurrent
            Exit Sub
        End If
        bases.datbasemenu.Recordset.Fields("tipoiva") = DataCombo1.BoundText
        bases.datbasemenu.Recordset.Fields("tipdocu") = Combo2.Text
        bases.datbasemenu.Recordset.Fields("cuit") = Text1(7).Text

        bases.datbasemenu.Recordset.Fields("zonafiscal") = Text1(8).Text
        If flagnuevo = 1 Then
              bases.datbasemenu.Recordset.Fields("codcliente") = Text1(9).Text
        End If
        If Text1(10).Text = "" Then Text1(10).Text = 0
        bases.datbasemenu.Recordset.Fields("codcontable") = Text1(10).Text
        If Text1(11).Text = "" Then Text1(11).Text = 0
        bases.datbasemenu.Recordset.Fields("codcontableventas") = Text1(11).Text
        bases.datbasemenu.Recordset.Fields("impcateg1") = Text1(12).Text
        bases.datbasemenu.Recordset.Fields("impnroincr1") = Text1(13).Text
        bases.datbasemenu.Recordset.Fields("impcateg2") = Text1(14).Text
        bases.datbasemenu.Recordset.Fields("impnroincr2") = Text1(15).Text
        bases.datbasemenu.Recordset.Fields("condventa") = Text1(16).Text
        bases.datbasemenu.Recordset.Fields("vendedor") = Text1(17).Text
        bases.datbasemenu.Recordset.Fields("cobrador") = Text1(18).Text
        bases.datbasemenu.Recordset.Fields("listaprecios") = Text1(19).Text
        bases.datbasemenu.Recordset.Fields("transportista") = Text1(20).Text
        If Combo1.Text = "Si" Or Combo1.Text = "" Then
            bases.datbasemenu.Recordset.Fields("activo") = 0
        Else
            bases.datbasemenu.Recordset.Fields("activo") = 1
        End If
        bases.datbasemenu.Recordset.Fields("tipocliente") = Text1(21).Text
        If Text1(22).Text = "" Then Text1(22).Text = "0.00"
        importe = Val(Text1(22).Text)
        bases.datbasemenu.Recordset.Fields("creditomax") = importe
        bases.datbasemenu.Recordset.Fields("tipobonif") = Text1(23).Text
        bases.datbasemenu.Recordset.Fields("motivo") = Text1(24).Text
        
        bases.datbasemenu.Recordset.UpdateBatch adAffectCurrent
        MsgBox "Almacenado Correctamente", vbInformation, "Guardar"
        lista_proveedores.lista = ""
        flagnuevo = 0
        Call llena_Click
        Text1(9).SetFocus

End Sub

Private Sub anterior_Click()
On Error Resume Next
    If bases.datbasemenu.Recordset.RecordCount = 1 Then
        bases.datbasemenu.RecordSource = "select clientes.* from clientes where empresa = " & empresareal & " order by razonsocial"
        bases.datbasemenu.Refresh
        bases.datbasemenu.Recordset.AbsolutePosition = lista_proveedores.posicion
        bases.datbasemenu.Recordset.MoveNext
    End If
    
    bases.datbasemenu.Recordset.MovePrevious

flagnuevo = 1
navega = 1
Call llena_Click
End Sub



Private Sub estcuenta_Click()

End Sub

Private Sub Cancelar_Click()

       lista_proveedores.lista = ""
       flagnuevo = 0
        Call llena_Click
        Text1(9).SetFocus
        
End Sub

Private Sub cerrar_Click()

    Unload Me

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(24).SetFocus
    End If

End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    KeyAscii = 0
    Text1(7).SetFocus
End If

End Sub

Private Sub Command1_GotFocus(Index As Integer)

    Text1(0).SetFocus

End Sub

Private Sub DataCombo1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        If DataCombo1.Text = "" Then
            MsgBox "Ingrese Condición Tributaria", vbCritical, "Error"
            DataCombo1.SetFocus
            Exit Sub
        End If
        DataGrid1.Bookmark = DataCombo1.SelectedItem
        Combo2.SetFocus
    End If

End Sub


Private Sub elminar_Click()
On Error Resume Next
    
    datverifica.RecordSource = "SELECT cuit, cliente, empresa From libroventas where empresa = " & login.empresaact & " GROUP BY cuit, cliente, empresa"
    datverifica.Refresh
    datverifica.Recordset.Filter = "cliente = '" & Text1(0).Text & "' and cuit = '" & Text1(7).Text & "'"
    If datverifica.Recordset.EOF = False Then
        MsgBox "Este Cliente tiene Comprobantes imputados, no se puede eliminar", vbCritical, "Error"
        Exit Sub
    End If
    
    datverifica.RecordSource = "SELECT codcliente, nomcliente, empresa From recibocobroabonan where empresa = " & login.empresaact & " GROUP BY codcliente, nomcliente, empresa"
    datverifica.Refresh
    datverifica.Recordset.Filter = "nomcliente = '" & Text1(0).Text & "' and codcliente = " & Text1(9).Text & ""
    If datverifica.Recordset.EOF = False Then
        MsgBox "Este Cliente tiene Recibos imputados, no se puede eliminar", vbCritical, "Error"
        Exit Sub
    End If


    mensa = MsgBox("Está por eliminar este registro, Esta Seguro ?", vbYesNo, "Atención")
    If mensa = vbYes Then
        bases.datbasemenu.Recordset.Delete adAffectCurrent
        Call Cancelar_Click
    End If
    

End Sub

Private Sub fechacai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1(15).SetFocus
    End If
End Sub

Private Sub Form_Load()
frmclientes.Top = 0
frmclientes.Left = 0

Aplicar_skin Me

If login.clientesaltas = "N" Or login.clientesmod = "N" Then
    aceptar.Enabled = False
Else
    aceptar.Enabled = True
End If

If login.clientesbajas = "N" Then
    elminar.Enabled = False
Else
    elminar.Enabled = True
End If


ventana.menu = 0
navega = 0
Combo2.AddItem "C.U.I.T."
Combo2.AddItem "D.N.I."
Combo2.AddItem "L.C."
Combo2.AddItem "L.E."
Combo2.AddItem "C.I."
Combo2.AddItem ""

Combo1.AddItem "Si"
Combo1.AddItem "No"

If login.empresaact1 > 0 Then
    empresareal = login.empresaact1
Else
    empresareal = login.empresaact
End If

datcondtrib.ConnectionString = login.conexiontotal
datverifica.ConnectionString = login.conexiontotal
datcondtrib.RecordSource = "select condtrib.* from condtrib"
datcondtrib.Refresh

bases.datbasemenu.ConnectionString = login.conexiontotal
bases.datbasemenu1.ConnectionString = login.conexiontotal
bases.datbasemenu.RecordSource = "select clientes.* from clientes where empresa = " & empresareal & " order by razonsocial"
bases.datbasemenu.Refresh
lista_proveedores.lista = ""

flagnuevo = 0



End Sub

Private Sub menaccion_Click()

End Sub

Private Sub KewlButtons1_Click()
On Error GoTo fuera

Dim tabla As String
Dim ruta As String

ruta = "\Empresa" + Right(Str(login.empresaact), Len(Str(login.empresaact)) - 1)

    mensa = MsgBox("Ordena por Codigo(Si), Razon Social(No)", vbYesNoCancel, "Listar")
    If mensa = vbYes Then reporte.SQL = "SELECT clientes.codcliente, clientes.razonsocial, clientes.tipoiva, clientes.cuit, clientes.domicilio, clientes.localidad, EMPRESA.razonsocial FROM { oj contablesql.dbo.clientes clientes INNER JOIN contablesql.dbo.EMPRESA EMPRESA ON clientes.empresa = EMPRESA.empresa} where clientes.empresa = " & empresareal & " ORDER BY clientes.codcliente ASC"
    If mensa = vbNo Then reporte.SQL = "SELECT clientes.codcliente, clientes.razonsocial, clientes.tipoiva, clientes.cuit, clientes.domicilio, clientes.localidad, EMPRESA.razonsocial FROM { oj contablesql.dbo.clientes clientes INNER JOIN contablesql.dbo.EMPRESA EMPRESA ON clientes.empresa = EMPRESA.empresa} where clientes.empresa = " & empresareal & " ORDER BY clientes.razonsocial ASC"
    If mensa = vbCancel Then Exit Sub

Rem reporte.SQL = "SELECT clientes.codcliente, clientes.razonsocial, clientes.tipoiva, clientes.cuit, clientes.domicilio, clientes.localidad, EMPRESA.razonsocial FROM { oj contablesql.dbo.clientes clientes INNER JOIN contablesql.dbo.EMPRESA EMPRESA ON clientes.empresa = EMPRESA.empresa} where clientes.empresa = " & login.empresaact & " ORDER BY clientes.codcliente ASC"
tabla = reporte.SQL

With CrystalReporte
    .ReportFileName = App.Path & ruta + "\clientes.rpt"
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

fuera:
End Sub

Private Sub llena_Click()
On Error Resume Next
        For x = 0 To 24
            Text1(x).Text = ""
        Next x
        
   If navega = 0 Then
        bases.datbasemenu.RecordSource = "select clientes.* from clientes where empresa = " & empresareal & " and razonsocial = '" & lista_proveedores.lista & "' "
        bases.datbasemenu.Refresh
        If bases.datbasemenu.Recordset.EOF = True Then Exit Sub
        Text1(0).Text = lista_proveedores.lista
   Else
        Text1(0).Text = bases.datbasemenu.Recordset.Fields("razonsocial")
   End If
        If IsNull(bases.datbasemenu.Recordset.Fields("domicilio")) = True Then
                Text1(1).Text = ""
        Else
                Text1(1).Text = bases.datbasemenu.Recordset.Fields("domicilio")
        End If
        If IsNull(bases.datbasemenu.Recordset.Fields("localidad")) = True Then
                Text1(2).Text = ""
        Else
                Text1(2).Text = bases.datbasemenu.Recordset.Fields("localidad")
        End If
        If IsNull(bases.datbasemenu.Recordset.Fields("codpostal")) Then
                Text1(3).Text = ""
        Else
                Text1(3).Text = bases.datbasemenu.Recordset.Fields("codpostal")
        End If
        If IsNull(bases.datbasemenu.Recordset.Fields("telefono")) Then
                Text1(4).Text = ""
        Else
                Text1(4).Text = bases.datbasemenu.Recordset.Fields("telefono")
        End If
        If IsNull(bases.datbasemenu.Recordset.Fields("email")) Then
                Text1(5).Text = ""
        Else
                Text1(5).Text = bases.datbasemenu.Recordset.Fields("email")
        End If
        If IsNull(bases.datbasemenu.Recordset.Fields("contacto")) Then
                Text1(6).Text = ""
        Else
                Text1(6).Text = bases.datbasemenu.Recordset.Fields("contacto")
        End If
        DataCombo1.BoundText = bases.datbasemenu.Recordset.Fields("tipoiva")
        If IsNull(bases.datbasemenu.Recordset.Fields("cuit")) Then
                Text1(7).Text = ""
        Else
                Text1(7).Text = bases.datbasemenu.Recordset.Fields("cuit")
        End If
        Text1(9).Text = bases.datbasemenu.Recordset.Fields("codcliente")
        If IsNull(bases.datbasemenu.Recordset.Fields("codcontable")) Then
                Text1(10).Text = ""
        Else
                Text1(10).Text = bases.datbasemenu.Recordset.Fields("codcontable")
        End If
        If IsNull(bases.datbasemenu.Recordset.Fields("codcontableventas")) Then
                Text1(11).Text = ""
        Else
                Text1(11).Text = bases.datbasemenu.Recordset.Fields("codcontableventas")
        End If
        If IsNull(bases.datbasemenu.Recordset.Fields("zonafiscal")) Then
                Text1(8).Text = ""
        Else
                Text1(8).Text = bases.datbasemenu.Recordset.Fields("zonafiscal")
        End If
        
        If IsNull(bases.datbasemenu.Recordset.Fields("tipdocu")) = False Then
                Combo2.Text = bases.datbasemenu.Recordset.Fields("tipdocu")
        Else
                Combo2.ListIndex = 5
        End If
        Text1(12).Text = bases.datbasemenu.Recordset.Fields("impcateg1")
        Text1(13).Text = bases.datbasemenu.Recordset.Fields("impnroincr1")
        Text1(14).Text = bases.datbasemenu.Recordset.Fields("impcateg2")
        Text1(15).Text = bases.datbasemenu.Recordset.Fields("impnroincr2")
        Text1(16).Text = bases.datbasemenu.Recordset.Fields("condventa")
        Text1(17).Text = bases.datbasemenu.Recordset.Fields("vendedor")
        Text1(18).Text = bases.datbasemenu.Recordset.Fields("cobrador")
        Text1(19).Text = bases.datbasemenu.Recordset.Fields("listaprecios")
        Text1(20).Text = bases.datbasemenu.Recordset.Fields("transportista")
        Text1(21).Text = bases.datbasemenu.Recordset.Fields("tipocliente")
        If IsNull(bases.datbasemenu.Recordset.Fields("activo")) = True Then
            bases.datbasemenu.Recordset.Fields("activo") = 0
        End If
        If bases.datbasemenu.Recordset.Fields("activo") = 0 Then
            Combo1.ListIndex = 0
        Else
            Combo1.ListIndex = 1
        End If
        If IsNull(bases.datbasemenu.Recordset.Fields("motivo")) = True Then
            Text1(24).Text = ""
        Else
            Text1(24).Text = bases.datbasemenu.Recordset.Fields("motivo")
        End If
        Text1(22).Text = bases.datbasemenu.Recordset.Fields("creditomax")
        Text1(22).Text = Format(Text1(22).Text, "#,##0.00")
        Text1(23).Text = bases.datbasemenu.Recordset.Fields("tipobonif")
        
        ventana.menu = 0
        navega = 0
        
End Sub

Private Sub salir_Click()

    Unload Me

End Sub

Private Sub siguiente_Click()
On Error Resume Next
    
    If bases.datbasemenu.Recordset.RecordCount = 1 Then
        bases.datbasemenu.RecordSource = "select clientes.* from clientes where empresa = " & empresareal & " order by razonsocial"
        bases.datbasemenu.Refresh
        bases.datbasemenu.Recordset.AbsolutePosition = lista_proveedores.posicion
        bases.datbasemenu.Recordset.MoveNext
    End If
    
    bases.datbasemenu.Recordset.MoveNext

flagnuevo = 1
navega = 1
Call llena_Click

End Sub

Private Sub Text1_Change(Index As Integer)

    If Index = 7 And Combo2.Text = "C.U.I.T." Then
        If Len(Text1(Index).Text) = 2 Or Len(Text1(Index).Text) = 11 Then
            Text1(Index).Text = Text1(Index).Text + "-"
            SendKeys "{end}", False
        End If
    End If

End Sub

Private Sub Text1_GotFocus(Index As Integer)

    If ventana.menu = 2 And (Index = 0 Or Index = 9) Then
        Call llena_Click
    End If
    

    If ventana.menu = 4 And Index = 10 Then
        ventana.menu = 0
        Text1(10).Text = lista_cuentas.cuentacont
    End If
    If ventana.menu = 4 And Index = 11 Then
        ventana.menu = 0
        Text1(11).Text = lista_cuentas.cuentacont
    End If
    If ventana.menu = 1 And Index = 16 Then
        ventana.menu = 0
        Text1(16).Text = lista_condventa.lista
    End If
    If ventana.menu = 1 And Index = 21 Then
        ventana.menu = 0
        Text1(21).Text = lista_tipoclientes.lista
    End If
    
    

    
    
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

    If Index = 7 And KeyAscii = 27 Then
            salir = 1
            Call Cancelar_Click
    End If

    If KeyAscii = 13 Then
        KeyAscii = 0
        salir = 0
                        
        If Index = 10 Or Index = 11 Then
            posicion = Index
            Call verificacuenta_Click
        End If
                       
        SendKeys "{tab}", False
        
    End If


End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 38 Then
        If Index = 0 Then
            Text1(9).SetFocus
            Exit Sub
        End If
        If Index = 9 Then Exit Sub
        Text1(Index - 1).SetFocus
        Exit Sub
    End If
    
    If KeyCode = 114 And (Index = 0 Or Index = 9) Then
        ventana.menu = 2
        lista_proveedores.Show
    End If
    
    If KeyCode = 117 And Index = 0 Then
        Call siguiente_Click
    End If
    If KeyCode = 116 And Index = 0 Then
        Call anterior_Click
    End If
    
    If KeyCode = 114 And Index = 16 Then
        ventana.menu = 1
        lista_condventa.Show
    End If
    
    If KeyCode = 114 And Index = 21 Then
        ventana.menu = 1
        lista_tipoclientes.Show
    End If
    
    If KeyCode = 114 And Index = 10 Then
        lista_cuentas.cuentacont = Text1(10).Text
        ventana.menu = 4
        lista_cuentas.Show
    End If
    If KeyCode = 114 And Index = 11 Then
        lista_cuentas.cuentacont = Text1(11).Text
        ventana.menu = 4
        lista_cuentas.Show
    End If
    
  
End Sub

Private Sub Text1_LostFocus(Index As Integer)
On Error Resume Next
Dim invalido As Integer

    If Index = 9 And ventana.menu = 0 And Text1(0).Text <> "" Then
        bases.datbasemenu1.RecordSource = "select codcliente,cuit,empresa,razonsocial from clientes"
        bases.datbasemenu1.Refresh
        bases.datbasemenu1.Recordset.Filter = "empresa = " & empresareal & " and codcliente = '" & Text1(9).Text & "'"
        If bases.datbasemenu1.Recordset.EOF = False Then
            lista_proveedores.lista = bases.datbasemenu1.Recordset.Fields("razonsocial")
            flagnuevo = 1
            Call llena_Click
        End If
    End If

    If Index = 7 And salir = 0 And DataGrid1.Columns(2).Text = "1" Then
            mensa = verifica_cuit(Text1(7).Text, invalido)
Rem ************  restringe el ingresode cuit si es invalido ********
Rem            If invalido = 1 Then
Rem                Text1(7).Text = ""
Rem                Text1(7).SetFocus
Rem            End If
            posicion = bases.datbasemenu.Recordset.AbsolutePosition
            bases.datbasemenu1.RecordSource = "select cuit,empresa,razonsocial from clientes where empresa = " & empresareal & " and cuit = '" & Text1(7).Text & "'"
            bases.datbasemenu1.Refresh
            If bases.datbasemenu1.Recordset.EOF = False Then
                If bases.datbasemenu1.Recordset.Fields("razonsocial") <> Text1(0).Text Then
                    mensa = MsgBox("Ya existe otro Cliente con el mismo Nro. de CUIT", vbCritical, "Error")
                    Text1(7).SetFocus
                End If
            End If
    End If
End Sub

Private Sub verificacuenta_Click()

    If Text1(posicion) = "" Then Exit Sub

    datcuentas.ConnectionString = login.conexiontotal
    datcuentas.RecordSource = "select cuentas.* from cuentas where empre = " & login.empresaact & " and imp = 'S' and [cod contable] = " & Text1(posicion).Text & " and inicioper = '" & login.iper & "'"
    datcuentas.Refresh
    
    If datcuentas.Recordset.EOF = True Then
        MsgBox "No Existe esta cuenta contable", vbCritical, "Verificar"
        Text1(posicion).Text = ""
        Text1(posicion).SetFocus
    End If
    

End Sub
