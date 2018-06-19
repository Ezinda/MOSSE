VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form frmconciliacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conciliación Bancaria"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8235
   ScaleWidth      =   14985
   Begin VB.Frame Frame6 
      Height          =   1455
      Left            =   10920
      TabIndex        =   41
      Top             =   240
      Width           =   3855
      Begin KewlButtonz.KewlButtons KewlButtons1 
         Height          =   615
         Left            =   240
         TabIndex        =   42
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
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
         MICON           =   "frmconciliacion.frx":0000
         PICN            =   "frmconciliacion.frx":001C
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
   Begin VB.Frame Frame5 
      Caption         =   "Mostrar"
      Height          =   1455
      Left            =   8880
      TabIndex        =   37
      Top             =   240
      Width           =   1935
      Begin VB.OptionButton Option9 
         Caption         =   "En Cartera o No Conciliado"
         Height          =   375
         Left            =   360
         TabIndex        =   40
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Cancelados"
         Height          =   255
         Left            =   360
         TabIndex        =   39
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Conciliados"
         Height          =   255
         Left            =   360
         TabIndex        =   38
         Top             =   240
         Width           =   1335
      End
   End
   Begin vbskpro.Skinner Skinner1 
      Left            =   13680
      Top             =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      OldForeColor    =   0
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
      LcK1            =   "3.66*/4/0*/1-5*210/."
      LcK2            =   $"frmconciliacion.frx":0A2E
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
   Begin VB.Frame Frame2 
      Height          =   6135
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   14535
      Begin KewlButtonz.KewlButtons corrigeconc 
         Height          =   735
         Left            =   13320
         TabIndex        =   43
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1296
         BTYPE           =   14
         TX              =   "Corregir"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15591915
         BCOLO           =   15591915
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmconciliacion.frx":0A3D
         PICN            =   "frmconciliacion.frx":0A59
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.PictureBox Picture1 
         Height          =   1095
         Left            =   10200
         ScaleHeight     =   1035
         ScaleWidth      =   75
         TabIndex        =   36
         Top             =   120
         Width           =   135
      End
      Begin VB.CommandButton Cuenta 
         Caption         =   "Fecha Conc."
         Height          =   255
         Index           =   7
         Left            =   10560
         Picture         =   "frmconciliacion.frx":146B
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   720
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   13560
         TabIndex        =   32
         Text            =   "Text6"
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   13560
         TabIndex        =   31
         Text            =   "Text2"
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cancminuta 
         Caption         =   "cancminuta"
         Height          =   375
         Left            =   6600
         TabIndex        =   30
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton Cuenta 
         Caption         =   "Saldo"
         Height          =   255
         Index           =   6
         Left            =   6120
         Picture         =   "frmconciliacion.frx":199D
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   5520
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton Cuenta 
         Caption         =   "Credito"
         Height          =   255
         Index           =   5
         Left            =   3840
         Picture         =   "frmconciliacion.frx":1ECF
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   5520
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton Cuenta 
         Caption         =   "Debito"
         Height          =   255
         Index           =   4
         Left            =   1560
         Picture         =   "frmconciliacion.frx":2401
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   5520
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton Cuenta 
         Caption         =   "Fec.Cont.      Fec.Cheque     Fec.Vencim         Nº Cheque                Debe                 Haber"
         Height          =   255
         Index           =   3
         Left            =   240
         Picture         =   "frmconciliacion.frx":2933
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   7215
      End
      Begin KewlButtonz.KewlButtons limpia 
         Height          =   495
         Left            =   1200
         TabIndex        =   19
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   14
         TX              =   "&Limpia Items"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15591915
         BCOLO           =   15591915
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmconciliacion.frx":2E65
         PICN            =   "frmconciliacion.frx":2E81
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Fecha"
         Height          =   195
         Left            =   3360
         TabIndex        =   16
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Nro."
         Height          =   195
         Left            =   2520
         TabIndex        =   15
         Top             =   480
         Width           =   735
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ordenear por"
         Height          =   615
         Left            =   2400
         TabIndex        =   11
         Top             =   240
         Width           =   4095
         Begin VB.Frame Frame4 
            Height          =   615
            Left            =   2160
            TabIndex        =   14
            Top             =   0
            Width           =   15
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Des"
            Height          =   255
            Left            =   3240
            TabIndex        =   13
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Asc"
            Height          =   255
            Left            =   2400
            TabIndex        =   12
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.ListBox List1 
         Height          =   3885
         ItemData        =   "frmconciliacion.frx":6273
         Left            =   240
         List            =   "frmconciliacion.frx":6275
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   1200
         Width           =   14055
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2400
         TabIndex        =   9
         Top             =   5520
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4680
         TabIndex        =   8
         Top             =   5520
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6960
         TabIndex        =   7
         Top             =   5520
         Width           =   1215
      End
      Begin KewlButtonz.KewlButtons asignartodos 
         Height          =   495
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         BTYPE           =   14
         TX              =   "&Tildar"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15591915
         BCOLO           =   15591915
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmconciliacion.frx":6277
         PICN            =   "frmconciliacion.frx":6293
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons cancelacion 
         Height          =   375
         Left            =   7680
         TabIndex        =   25
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "&Cancelación Directa"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15591915
         BCOLO           =   15591915
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmconciliacion.frx":6CA5
         PICN            =   "frmconciliacion.frx":6CC1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons cancelminuta 
         Height          =   375
         Left            =   7680
         TabIndex        =   26
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "Cancelación x &Minutas"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15591915
         BCOLO           =   15591915
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmconciliacion.frx":76D3
         PICN            =   "frmconciliacion.frx":76EF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons conciliar 
         Height          =   375
         Left            =   10560
         TabIndex        =   33
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "Conciliación &Bancaria"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15591915
         BCOLO           =   15591915
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmconciliacion.frx":7DC1
         PICN            =   "frmconciliacion.frx":7DDD
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker fechaconc 
         Height          =   285
         Left            =   11880
         TabIndex        =   34
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   64290817
         CurrentDate     =   39247
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8535
      Begin VB.CommandButton Cuenta 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   2
         Left            =   2520
         Picture         =   "frmconciliacion.frx":87EF
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   840
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton Cuenta 
         Caption         =   "Desde"
         Height          =   255
         Index           =   1
         Left            =   120
         Picture         =   "frmconciliacion.frx":8D21
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   840
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton Cuenta 
         Caption         =   "&Cuenta"
         Height          =   255
         Index           =   0
         Left            =   120
         Picture         =   "frmconciliacion.frx":9253
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option6 
         Caption         =   "x Fecha Vencimiento"
         Height          =   195
         Left            =   5160
         TabIndex        =   18
         Top             =   1080
         Width           =   2055
      End
      Begin VB.OptionButton Option5 
         Caption         =   "x Fecha Contable"
         Height          =   195
         Left            =   5160
         TabIndex        =   17
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton aceptar 
         Caption         =   "&Aceptar"
         Height          =   735
         Left            =   7440
         Picture         =   "frmconciliacion.frx":9785
         TabIndex        =   5
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker fechadesde 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   64290817
         CurrentDate     =   39247
      End
      Begin MSComCtl2.DTPicker fechahasta 
         Height          =   285
         Left            =   3480
         TabIndex        =   2
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   64290817
         CurrentDate     =   39247
      End
      Begin MSDataListLib.DataCombo DataCombo4 
         Bindings        =   "frmconciliacion.frx":9CB7
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Nombre Cuenta"
         BoundColumn     =   "cuenta"
         Text            =   "DataCombo4"
      End
   End
   Begin MSAdodcLib.Adodc datconciliacion 
      Height          =   330
      Left            =   960
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
   Begin MSAdodcLib.Adodc datcheques 
      Height          =   330
      Left            =   2280
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
   Begin MSAdodcLib.Adodc datcomprob 
      Height          =   330
      Left            =   3480
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select conciliacioncomprob.* from conciliacioncomprob"
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
   Begin MSAdodcLib.Adodc datfondocaja 
      Height          =   330
      Left            =   4680
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
   Begin MSAdodcLib.Adodc datempresa 
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
   Begin MSAdodcLib.Adodc datinstrumento 
      Height          =   330
      Left            =   6000
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
   Begin MSAdodcLib.Adodc datcancelacion 
      Height          =   330
      Left            =   7320
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
   Begin Crystal.CrystalReport CrystalReporte 
      Left            =   8520
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Listado de Recibos Emitidos"
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSRDC.MSRDC reporte 
      Height          =   375
      Left            =   8880
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
   Begin MSAdodcLib.Adodc criterio 
      Height          =   330
      Left            =   0
      Top             =   240
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
Attribute VB_Name = "frmconciliacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tipo(99999) As String
Dim idinstru(99999) As Double
Dim cancelado(99999) As Boolean
Dim campoasiento(99999) As Double
Dim camponroasiento(99999) As Double
Dim campoperiodo(99999) As Double
Dim as_cuenta(99999) As Integer
Dim as_debe(99999) As Currency
Dim as_haber(99999) As Currency
Dim as_desc(99999) As String
Dim as_cc(99999) As Integer
Public lineagrilla As Integer

Private Function chequea(ByVal Numero As Integer)
On Error Resume Next

       If List1.Selected(Numero) = True Then
            Text3.Text = Text3.Text + datcheques.Recordset.Fields("ingreso")
            Text4.Text = Text4.Text + datcheques.Recordset.Fields("egreso")
       Else
            Text3.Text = Text3.Text - datcheques.Recordset.Fields("ingreso")
            Text4.Text = Text4.Text - datcheques.Recordset.Fields("egreso")
       End If
       Text5.Text = Text3.Text - Text4.Text
       Text3.Text = Format(Text3.Text, "#,##0.00")
       Text4.Text = Format(Text4.Text, "#,##0.00")
       Text5.Text = Format(Text5.Text, "#,##0.00")


End Function

Private Sub aceptar_Click()

Call llena_Click

End Sub

Private Sub asignartodos_Click()

datcheques.Recordset.MoveFirst

Text3.Text = 0
Text4.Text = 0
Text5.Text = 0

i = 0
Do While Not datcheques.Recordset.EOF
     
     List1.Selected(i) = True
     datcheques.Recordset.MoveNext
     i = i + 1

Loop

For i = 0 To i - 1
         chequea (i)
Next i

End Sub

Private Sub Command1_Click()

    datcomprob.Recordset.AddNew
    datcomprob.Recordset.Fields("nro") = text1.Text
    datcomprob.Recordset.Fields("empresa") = login.empresaact
    


End Sub

Private Sub Command2_Click()

    datcomprob.Recordset.Delete adAffectCurrent
    datcomprob.Refresh
    
End Sub





Private Sub conciliar_Click()

    datcancelacion.RecordSource = "select cancelacion.* from cancelacion where empresa = " & login.empresaact & " and idinstru = '0' "
    datcancelacion.Refresh
        
    For x = 0 To List1.ListCount - 1
         If idinstru(x) <> 0 Then
          If tipo(x) = "O" And List1.Selected(x) = True Then
            datinstrumento.RecordSource = "select ordendepagoinstrumento.* from ordendepagoinstrumento where empresa = " & login.empresaact & " and id = " & idinstru(x) & " "
            datinstrumento.Refresh
            If datinstrumento.Recordset.EOF = False Then
                datinstrumento.Recordset.Fields("conciliado") = List1.Selected(x)
                datinstrumento.Recordset.Fields("conciliacion") = List1.Selected(x)
                datinstrumento.Recordset.Fields("fechaconciliacion") = fechaconc.Value
                datinstrumento.Recordset.UpdateBatch adAffectCurrent
            End If
          End If
          If (tipo(x) = "RB" Or tipo(x) = "OB") And List1.Selected(x) = True Then
            datinstrumento.RecordSource = "select librocajabanco.* from librocajabanco where empresa = " & login.empresaact & " and id = " & idinstru(x) & " "
            datinstrumento.Refresh
            If datinstrumento.Recordset.EOF = False Then
                datinstrumento.Recordset.Fields("conciliado") = List1.Selected(x)
                datinstrumento.Recordset.Fields("conciliacion") = List1.Selected(x)
                datinstrumento.Recordset.Fields("fechaconciliacion") = fechaconc.Value
                datinstrumento.Recordset.UpdateBatch adAffectCurrent
            End If
          End If
          
         End If
    Next x
    

    Call llena_Click
End Sub

Private Sub cancelacion_Click()
On Error GoTo fuera

contar = 0
For x = 0 To List1.ListCount - 1
 If List1.Selected(x) = True Then contar = contar + 1
Next x

If contar > 2 Then
        MsgBox "No puede seleccionar mas de dos Items a Cancelar", vbCritical, "Error"
        List1.SetFocus
        Exit Sub
End If

    If Text5.Text <> 0 Then
        MsgBox "El saldo de los comprobantes a cancelar debe ser cero", vbCritical, "Error"
        List1.SetFocus
        Exit Sub
    End If
        
    datcancelacion.RecordSource = "select cancelacion.* from cancelacion where empresa = " & login.empresaact & " and idinstru = '0' "
    datcancelacion.Refresh
    asientopri = 0
    asientonropri = 0
    movimiento = 0
    For x = 0 To List1.ListCount - 1
          If idinstru(x) <> 0 Then
            If tipo(x) = "R" And List1.Selected(x) = True Then
                datinstrumento.RecordSource = "select recibocobroinstrumento.* from recibocobroinstrumento where empresa = " & login.empresaact & " and id = " & idinstru(x) & " "
                datinstrumento.Refresh
                movimiento = movimiento + 1
                GoTo sigue0
            End If
            If tipo(x) = "O" And List1.Selected(x) = True Then
                datinstrumento.RecordSource = "select ordendepagoinstrumento.* from ordendepagoinstrumento where empresa = " & login.empresaact & " and id = " & idinstru(x) & " "
                datinstrumento.Refresh
                movimiento = movimiento + 1
                GoTo sigue0
            End If
          End If
          If List1.Selected(x) = True Then
            asientopri = campoasiento(x)
            asientonropri = camponroasiento(x)
            asientoperiodo = campoperiodo(x)
            movimiento = movimiento + 1
          End If

          GoTo sigue01

sigue0:
         If datinstrumento.Recordset.EOF = False Then
                datinstrumento.Recordset.Fields("conciliado") = List1.Selected(x)
                datinstrumento.Recordset.UpdateBatch adAffectCurrent
         End If
         campocheque0 = tipo(x)
         campocheque1 = idinstru(x)
         campocheque2 = datinstrumento.Recordset.Fields("instrumento")
         campocheque3 = datinstrumento.Recordset.Fields("denominacion")
         campocheque4 = datinstrumento.Recordset.Fields("comprobante")
         campocheque5 = datinstrumento.Recordset.Fields("fechacompro")
         campocheque6 = datinstrumento.Recordset.Fields("fechavencim")
         campocheque7 = datinstrumento.Recordset.Fields("importe")
         campocheque8 = datinstrumento.Recordset.Fields("codcuenta")
sigue01:
         If movimiento = 1 Then GoTo sigue1
         
         If List1.Selected(x) = True Then
                datcancelacion.Recordset.AddNew
                datcancelacion.Recordset.Fields("tipo") = campocheque0
                datcancelacion.Recordset.Fields("idinstru") = campocheque1
                If asientopri = 0 Then
                    datcancelacion.Recordset.Fields("idasiento") = campoasiento(x)
                    datcancelacion.Recordset.Fields("asiento") = camponroasiento(x)
                    datcancelacion.Recordset.Fields("inicioper") = campoperiodo(x)
                Else
                    datcancelacion.Recordset.Fields("idasiento") = asientopri
                    datcancelacion.Recordset.Fields("asiento") = asientonropri
                    datcancelacion.Recordset.Fields("inicioper") = asientoperiodo
                End If
                datcancelacion.Recordset.Fields("instrumento") = campocheque2
                datcancelacion.Recordset.Fields("denominacion") = campocheque3
                datcancelacion.Recordset.Fields("comprobante") = campocheque4
                datcancelacion.Recordset.Fields("fechacompro") = campocheque5
                datcancelacion.Recordset.Fields("fechavencim") = campocheque6
                datcancelacion.Recordset.Fields("importe") = campocheque7
                datcancelacion.Recordset.Fields("codcuenta") = campocheque8
                datcancelacion.Recordset.Fields("conciliado") = True
                datcancelacion.Recordset.Fields("empresa") = login.empresaact
                datcancelacion.Recordset.Fields("fechacancel") = Date
                datcancelacion.Recordset.UpdateBatch
         End If
         
    
sigue1:
    Next x

    Call llena_Click
    Exit Sub
fuera:
    MsgBox "Error, esta instando duplicar una cancelacion", vbCritical, "Error"

End Sub

Private Sub cancelminuta_Click()

Y = 1
frmasientoscancelacion.MaskEdBox1 = Date


For x = 0 To List1.ListCount - 1
    If List1.Selected(x) = True Then
        frmasientoscancelacion.grilla.Row = Y
        frmasientoscancelacion.grilla.Col = 1
        frmasientoscancelacion.grilla.Text = as_cuenta(x + 1)
        frmasientoscancelacion.grilla.Col = 2
        frmasientoscancelacion.grilla.Text = as_debe(x + 1)
        frmasientoscancelacion.grilla.Col = 3
        frmasientoscancelacion.grilla.Text = as_haber(x + 1)
        frmasientoscancelacion.grilla.Col = 4
        frmasientoscancelacion.grilla.Text = as_desc(x + 1)
        frmasientoscancelacion.grilla.Col = 5
        frmasientoscancelacion.grilla.Text = as_cc(x + 1)
        Y = Y + 1
    End If
Next x
           
sumadebe = 0
sumahaber = 0
For x = 1 To Y
    frmasientoscancelacion.grilla.Row = x
    frmasientoscancelacion.grilla.Col = 2
    frmasientoscancelacion.Text6(0).Text = Replace(frmasientoscancelacion.grilla.Text, ",", "")
    sumadebe = Val(frmasientoscancelacion.Text6(0).Text) + sumadebe
    frmasientoscancelacion.grilla.Col = 3
    frmasientoscancelacion.Text6(1).Text = Replace(frmasientoscancelacion.grilla.Text, ",", "")
    sumahaber = Val(frmasientoscancelacion.Text6(1).Text) + sumahaber
 Next x

    frmasientoscancelacion.Maskdebe.Text = sumadebe
    frmasientoscancelacion.Maskhaber.Text = sumahaber
    frmasientoscancelacion.Masksaldo.Text = sumadebe - sumahaber

frmasientoscancelacion.Show


End Sub

Private Sub cancelminuta_GotFocus()

    If lineagrilla = 1 Then Call cancminuta_Click

End Sub

Private Sub cancminuta_Click()

On Error GoTo fuera
    grillalinea = 0
    
    datcancelacion.RecordSource = "select cancelacion.* from cancelacion where empresa = " & login.empresaact & " and idinstru = '0' "
    datcancelacion.Refresh
    Y = 0
    For x = 0 To List1.ListCount - 1
         If idinstru(x) <> 0 Then
          If tipo(x) = "R" And List1.Selected(x) = True Then
            datinstrumento.RecordSource = "select recibocobroinstrumento.* from recibocobroinstrumento where empresa = " & login.empresaact & " and id = " & idinstru(x) & " "
            datinstrumento.Refresh
            GoTo sigue0
          End If
          If tipo(x) = "O" And List1.Selected(x) = True Then
            datinstrumento.RecordSource = "select ordendepagoinstrumento.* from ordendepagoinstrumento where empresa = " & login.empresaact & " and id = " & idinstru(x) & " "
            datinstrumento.Refresh
            GoTo sigue0
          End If
          GoTo sigue1
sigue0:
          If datinstrumento.Recordset.EOF = False Then
                datinstrumento.Recordset.Fields("conciliado") = List1.Selected(x)
                datinstrumento.Recordset.UpdateBatch adAffectCurrent
          End If
          campocheque0 = tipo(x)
          campocheque1 = idinstru(x)
          campocheque2 = datinstrumento.Recordset.Fields("instrumento")
          campocheque3 = datinstrumento.Recordset.Fields("denominacion")
          campocheque4 = datinstrumento.Recordset.Fields("comprobante")
          campocheque5 = datinstrumento.Recordset.Fields("fechacompro")
          campocheque6 = datinstrumento.Recordset.Fields("fechavencim")
          campocheque7 = datinstrumento.Recordset.Fields("importe")
          campocheque8 = datinstrumento.Recordset.Fields("codcuenta")
          If List1.Selected(x) = True Then
                datcancelacion.Recordset.AddNew
                datcancelacion.Recordset.Fields("tipo") = campocheque0
                datcancelacion.Recordset.Fields("idinstru") = campocheque1
                datcancelacion.Recordset.Fields("idasiento") = Text2.Text + Y
                datcancelacion.Recordset.Fields("asiento") = Text6.Text
                datcancelacion.Recordset.Fields("inicioper") = login.iper
                datcancelacion.Recordset.Fields("instrumento") = campocheque2
                datcancelacion.Recordset.Fields("denominacion") = campocheque3
                datcancelacion.Recordset.Fields("comprobante") = campocheque4
                datcancelacion.Recordset.Fields("fechacompro") = campocheque5
                datcancelacion.Recordset.Fields("fechavencim") = campocheque6
                datcancelacion.Recordset.Fields("importe") = campocheque7
                datcancelacion.Recordset.Fields("codcuenta") = campocheque8
                datcancelacion.Recordset.Fields("conciliado") = True
                datcancelacion.Recordset.Fields("empresa") = login.empresaact
                datcancelacion.Recordset.Fields("fechacancel") = Date
                Y = Y + 1
                datcancelacion.Recordset.UpdateBatch
          End If
                                     
        End If
               
sigue1:
    Next x
    cuentaprovi = DataCombo4.Text
    cuentaprovi1 = DataCombo4.BoundText
    fechadesde1 = fechadesde.Value
    fechahasta1 = fechahasta.Value
    Unload Me
    frmconciliacion.Show
    DataCombo4.Text = cuentaprovi
    DataCombo4.BoundText = cuentaprovi1
    text1.Text = cuentaprovi1
    fechadesde.Value = fechadesde1
    fechahasta.Value = fechahasta1
    Call llena_Click
   
    
Exit Sub

fuera:

    
End Sub

Private Sub corrigeconc_Click()
    datcancelacion.RecordSource = "select cancelacion.* from cancelacion where empresa = " & login.empresaact & " and idinstru = '0' "
    datcancelacion.Refresh
        
    For x = 0 To List1.ListCount - 1
         If idinstru(x) <> 0 Then
          If tipo(x) = "O" And List1.Selected(x) = False Then
            datinstrumento.RecordSource = "select ordendepagoinstrumento.* from ordendepagoinstrumento where empresa = " & login.empresaact & " and id = " & idinstru(x) & " "
            datinstrumento.Refresh
            If datinstrumento.Recordset.EOF = False Then
                datinstrumento.Recordset.Fields("conciliado") = List1.Selected(x)
                datinstrumento.Recordset.Fields("conciliacion") = List1.Selected(x)
                datinstrumento.Recordset.Fields("fechaconciliacion") = Null
                datinstrumento.Recordset.UpdateBatch adAffectCurrent
            End If
          End If
         
          If (tipo(x) = "RB" Or tipo(x) = "OB") And List1.Selected(x) = True Then
            datinstrumento.RecordSource = "select librocajabanco.* from librocajabanco where empresa = " & login.empresaact & " and id = " & idinstru(x) & " "
            datinstrumento.Refresh
            If datinstrumento.Recordset.EOF = False Then
                datinstrumento.Recordset.Fields("conciliado") = List1.Selected(x)
                datinstrumento.Recordset.Fields("conciliacion") = List1.Selected(x)
                datinstrumento.Recordset.Fields("fechaconciliacion") = Null
                datinstrumento.Recordset.UpdateBatch adAffectCurrent
            End If
          End If
         
         End If
    Next x
    

    Call llena_Click
End Sub

Private Sub DataCombo4_Click(Area As Integer)

    text1.Text = DataCombo4.BoundText

End Sub

Private Sub DataCombo4_KeyUp(KeyCode As Integer, Shift As Integer)
    text1.Text = DataCombo4.BoundText
End Sub

Private Sub Form_Activate()

Rem  Option8.Enabled = True
Rem  Option9.Enabled = True
Rem  Option7.Enabled = True
Rem  Option1.Enabled = True
Rem  Option2.Enabled = True
Rem  Option3.Enabled = True
Rem  Option4.Enabled = True

End Sub

Private Sub Form_Load()
On Error GoTo fuera
Aplicar_skin Me

datconciliacion.ConnectionString = login.conexiontotal
datcheques.ConnectionString = login.conexiontotal
datfondocaja.ConnectionString = login.conexiontotal
datempresa.ConnectionString = login.conexiontotal
datinstrumento.ConnectionString = login.conexiontotal
datcancelacion.ConnectionString = login.conexiontotal


frmconciliacion.Top = 0
frmconciliacion.Left = 0


Option1.Value = True
Option3.Value = True
Option5.Value = True
Option9.Value = True
fechadesde.Value = Date - 30
fechahasta.Value = Date
fechaconc.Value = Date

datempresa.RecordSource = "select empresa.* from empresa where empresa = '" & login.empresaact & "'"
datempresa.Refresh

datfondocaja.RecordSource = "select fondos.* from fondos where empresa = " & login.empresaact & " and id = " & datempresa.Recordset.Fields("cuentalibrobanco") & " and prorrateo > 0 "
datfondocaja.Refresh
lineagrilla = 0

Exit Sub

fuera:
MsgBox "No Estan parametrizados los Fondos", vbCritical, "Error"
Unload Me

End Sub

Private Sub Label6_Click()
End Sub

Private Sub KewlButtons1_Click()
On Error GoTo fuera
Dim tabla As String
Dim tabla1 As String
Dim desdeprov As String
Dim hastaprov As String
Dim ruta As String
Dim reportever As String
Dim saldomuestra As String


        ruta = "\Empresa" + Right(Str(login.empresaact), Len(Str(login.empresaact)) - 1)
        
Rem   *************  Muestra En cartera y No conciliados  **********************

If Option1.Value = True Then
    If Option3.Value = True Then
       If Option5.Value = True Then
            reporte.SQL = "SELECT libro_cheque.Fecha, libro_cheque.concepto, libro_cheque.detalle, libro_cheque.ingreso, libro_cheque.egreso, libro_cheque.codcuenta, libro_cheque.detallecuenta, libro_cheque.comprobante, libro_cheque.fechavencim, libro_cheque.conciliado, libro_cheque.fechaconciliacion FROM contablesql.dbo.libro_cheque libro_cheque where libro_cheque.codcuenta = " & Val(DataCombo4.BoundText) & " and libro_cheque.fecha >= '" & fechadesde.Value & "' and libro_cheque.fecha <= '" & fechahasta.Value & "' order by libro_cheque.fecha"
       Else
            reporte.SQL = "SELECT libro_cheque.Fecha, libro_cheque.concepto, libro_cheque.detalle, libro_cheque.ingreso, libro_cheque.egreso, libro_cheque.codcuenta, libro_cheque.detallecuenta, libro_cheque.comprobante, libro_cheque.fechavencim, libro_cheque.conciliado, libro_cheque.fechaconciliacion FROM contablesql.dbo.libro_cheque libro_cheque where libro_cheque.codcuenta = " & Val(DataCombo4.BoundText) & " and libro_cheque.fechavencim >= '" & fechadesde.Value & "' and libro_cheque.fechavencim <= '" & fechahasta.Value & "' order by libro_cheque.fechavencim"
       End If
    Else
       If Option5.Value = True Then
            reporte.SQL = "SELECT libro_cheque.Fecha, libro_cheque.concepto, libro_cheque.detalle, libro_cheque.ingreso, libro_cheque.egreso, libro_cheque.codcuenta, libro_cheque.detallecuenta, libro_cheque.comprobante, libro_cheque.fechavencim, libro_cheque.conciliado, libro_cheque.fechaconciliacion FROM contablesql.dbo.libro_cheque libro_cheque where libro_cheque.codcuenta = " & Val(DataCombo4.BoundText) & " and libro_cheque.fecha >= '" & fechadesde.Value & "' and libro_cheque.fecha <= '" & fechahasta.Value & "' order by libro_cheque.fecha desc"
       Else
            reporte.SQL = "SELECT libro_cheque.Fecha, libro_cheque.concepto, libro_cheque.detalle, libro_cheque.ingreso, libro_cheque.egreso, libro_cheque.codcuenta, libro_cheque.detallecuenta, libro_cheque.comprobante, libro_cheque.fechavencim, libro_cheque.conciliado, libro_cheque.fechaconciliacion FROM contablesql.dbo.libro_cheque libro_cheque where libro_cheque.codcuenta = " & Val(DataCombo4.BoundText) & " and libro_cheque.fechavencim >= '" & fechadesde.Value & "' and libro_cheque.fechavencim <= '" & fechahasta.Value & "' order by libro_cheque.fechavencim desc"
       End If
    End If
End If

If Option2.Value = True Then
    If Option3.Value = True Then
       If Option5.Value = True Then
            reporte.SQL = "SELECT libro_cheque.Fecha, libro_cheque.concepto, libro_cheque.detalle, libro_cheque.ingreso, libro_cheque.egreso, libro_cheque.codcuenta, libro_cheque.detallecuenta, libro_cheque.comprobante, libro_cheque.fechavencim, libro_cheque.conciliado, libro_cheque.fechaconciliacion FROM contablesql.dbo.libro_cheque libro_cheque where libro_cheque.codcuenta = " & Val(DataCombo4.BoundText) & " and libro_cheque.fecha >= '" & fechadesde.Value & "' and libro_cheque.fecha <= '" & fechahasta.Value & "' order by libro_cheque.comprobante"
       Else
            reporte.SQL = "SELECT libro_cheque.Fecha, libro_cheque.concepto, libro_cheque.detalle, libro_cheque.ingreso, libro_cheque.egreso, libro_cheque.codcuenta, libro_cheque.detallecuenta, libro_cheque.comprobante, libro_cheque.fechavencim, libro_cheque.conciliado, libro_cheque.fechaconciliacion FROM contablesql.dbo.libro_cheque libro_cheque where libro_cheque.codcuenta = " & Val(DataCombo4.BoundText) & " and libro_cheque.fechavencim >= '" & fechadesde.Value & "' and libro_cheque.fechavencim <= '" & fechahasta.Value & "' order by libro_cheque.comprobante"
       End If
    Else
       If Option5.Value = True Then
            reporte.SQL = "SELECT libro_cheque.Fecha, libro_cheque.concepto, libro_cheque.detalle, libro_cheque.ingreso, libro_cheque.egreso, libro_cheque.codcuenta, libro_cheque.detallecuenta, libro_cheque.comprobante, libro_cheque.fechavencim, libro_cheque.conciliado, libro_cheque.fechaconciliacion FROM contablesql.dbo.libro_cheque libro_cheque where libro_cheque.codcuenta = " & Val(DataCombo4.BoundText) & " and libro_cheque.fecha >= '" & fechadesde.Value & "' and libro_cheque.fecha <= '" & fechahasta.Value & "' order by libro_cheque.comprobante desc"
       Else
            reporte.SQL = "SELECT libro_cheque.Fecha, libro_cheque.concepto, libro_cheque.detalle, libro_cheque.ingreso, libro_cheque.egreso, libro_cheque.codcuenta, libro_cheque.detallecuenta, libro_cheque.comprobante, libro_cheque.fechavencim, libro_cheque.conciliado, libro_cheque.fechaconciliacion FROM contablesql.dbo.libro_cheque libro_cheque where libro_cheque.codcuenta = " & Val(DataCombo4.BoundText) & " and libro_cheque.fechavencim >= '" & fechadesde.Value & "' and libro_cheque.fechavencim <= '" & fechahasta.Value & "' order by libro_cheque.comprobante desc"
       End If
    End If
End If


tabla = reporte.SQL

With CrystalReporte

If Option9.Value = True Then .ReportFileName = App.Path & ruta + "\cheques.rpt"
If Option8.Value = True Or Option7.Value = True Then .ReportFileName = App.Path & ruta + "\cheques_canc.rpt"

    .Connect = login.conexionreporte
    .Formulas(0) = "desdefecha=""" & fechadesde.Value & """"
    .Formulas(1) = "hastafecha=""" & fechahasta.Value & """"
    .Formulas(2) = "dempresa=""" & login.razonsoc & """"
Rem    .Formulas(1) = "hastafecha=""" & cargahasta.Value & """"
Rem    .Formulas(2) = "empresa=""" & login.nomempresa & """"
Rem    .Formulas(3) = "saldocero=""" & saldomuestra & """"
Rem    .SubreportToChange = .GetNthSubreportName(0)
Rem    .Connect = login.conexionreporte
Rem    .SubreportToChange = ""
Rem    .Connect = login.conexionreporte
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

Private Sub limpia_Click()

i = 0

Do While Not i = List1.ListCount
   
   List1.Selected(i) = False
   i = i + 1

Loop
   Text3.Text = 0
   Text4.Text = 0
   Text5.Text = 0

End Sub

Private Sub List1_ItemCheck(Item As Integer)
        
        
       datcheques.Recordset.AbsolutePosition = List1.ListIndex + 1
       chequea (List1.ListIndex)
      
       

End Sub

Private Sub llena_Click()
On Error Resume Next

List1.Clear
Text3.Text = 0
Text4.Text = 0
Text5.Text = 0

criterio.ConnectionString = login.conexiontotal
    
  criterio.RecordSource = "select empreactiva.* from empreactiva"
  criterio.Refresh

criterio.Recordset.Fields("empresa") = login.empresaact
criterio.Recordset.Fields("cuentacheque") = DataCombo4.BoundText
criterio.Recordset.UpdateBatch adAffectCurrent


If Option1.Value = True Then
    If Option3.Value = True Then
      If Option5.Value = True Then
        datcheques.RecordSource = "select libro_cheque.* from libro_cheque where empresa = " & login.empresaact & " and fecha >= '" & fechadesde.Value & "' and fecha <= '" & fechahasta.Value & "' order by fecha"
        datcheques.Refresh
      Else
        datcheques.RecordSource = "select libro_cheque.* from libro_cheque where empresa = " & login.empresaact & " and fechavencim >= '" & fechadesde.Value & "' and fechavencim <= '" & fechahasta.Value & "' order by fechavencim"
        datcheques.Refresh
      End If
    Else
      If Option5.Value = True Then
        datcheques.RecordSource = "select libro_cheque.* from libro_cheque where empresa = " & login.empresaact & " and fecha >= '" & fechadesde.Value & "' and fecha <= '" & fechahasta.Value & "' order by fecha desc"
        datcheques.Refresh
      Else
        datcheques.RecordSource = "select libro_cheque.* from libro_cheque where empresa = " & login.empresaact & " and fechavencim >= '" & fechadesde.Value & "' and fechavencim <= '" & fechahasta.Value & "' order by fechavencim desc"
        datcheques.Refresh
      End If
    End If
End If



If Option2.Value = True Then
    If Option3.Value = True Then
      If Option5.Value = True Then
        datcheques.RecordSource = "select libro_cheque.* from libro_cheque where empresa = " & login.empresaact & " and codcuenta = '" & DataCombo4.BoundText & "' and fecha >= '" & fechadesde.Value & "' and fecha <= '" & fechahasta.Value & "' order by comprobante"
        datcheques.Refresh
      Else
        datcheques.RecordSource = "select libro_cheque.* from libro_cheque where empresa = " & login.empresaact & " and codcuenta = '" & DataCombo4.BoundText & "' and fechavencim >= '" & fechadesde.Value & "' and fechavencim <= '" & fechahasta.Value & "' order by comprobante"
        datcheques.Refresh
      End If
    Else
      If Option5.Value = True Then
        datcheques.RecordSource = "select libro_cheque.* from libro_cheque where empresa = " & login.empresaact & " and codcuenta = '" & DataCombo4.BoundText & "' and fecha >= '" & fechadesde.Value & "' and fecha <= '" & fechahasta.Value & "' order by comprobante desc"
        datcheques.Refresh
      Else
        datcheques.RecordSource = "select libro_cheque.* from libro_cheque where empresa = " & login.empresaact & " and codcuenta = '" & DataCombo4.BoundText & "' and fechavencim >= '" & fechadesde.Value & "' and fechavencim <= '" & fechahasta.Value & "' order by comprobante desc"
        datcheques.Refresh
      End If
    End If
End If


If Option9.Value = True Then
    datcheques.Recordset.Filter = "conciliado = '0' or conciliado = NULL "
    cancelacion.Enabled = True
    cancelminuta.Enabled = True
    conciliar.Enabled = True
    corrigeconc.Enabled = False
Else
    cancelacion.Enabled = False
    cancelminuta.Enabled = False
    conciliar.Enabled = False
End If

If Option8.Value = True Then
    datcheques.Recordset.Filter = "conciliado =  True and conciliacion = Null"
    corrigeconc.Enabled = False
End If
If Option7.Value = True Then
    datcheques.Recordset.Filter = "conciliacion = '1' "
    corrigeconc.Enabled = True
End If

    


If datcheques.Recordset.EOF = True Then Exit Sub

datcheques.Recordset.MoveFirst
i = 0

Do While Not datcheques.Recordset.EOF
           
     i = i + 1
        
     campo1 = datcheques.Recordset.Fields("fecha")
     campo1a = datcheques.Recordset.Fields("fechacompro")
     campo1b = datcheques.Recordset.Fields("fechavencim")
        
 

     campo2 = datcheques.Recordset.Fields("comprobante")
     If IsNull(campo2) = True Then campo2 = "x"
     
     as_cuenta(i) = datcheques.Recordset.Fields("codcuenta")
     as_debe(i) = datcheques.Recordset.Fields("egreso")
     as_haber(i) = datcheques.Recordset.Fields("ingreso")
     as_desc(i) = campo2
     If IsNull(datcheques.Recordset.Fields("ccosto")) = False Then
        as_cc(i) = datcheques.Recordset.Fields("ccosto")
     Else
        as_cc(i) = 0
     End If
     
     campo3 = datcheques.Recordset.Fields("concepto")
     campo4 = datcheques.Recordset.Fields("ingreso")
     campo5 = datcheques.Recordset.Fields("egreso")
     campo6 = datcheques.Recordset.Fields("detalle")
     tipo(i - 1) = datcheques.Recordset.Fields("tipo")
     If IsNull(datcheques.Recordset.Fields("idinstrumento")) = True Then
            idinstru(i - 1) = 0
     Else
            idinstru(i - 1) = datcheques.Recordset.Fields("idinstrumento")
     End If
     If IsNull(datcheques.Recordset.Fields("conciliado")) = True Then
            cancelado(i - 1) = False
     Else
            cancelado(i - 1) = datcheques.Recordset.Fields("conciliado")
     End If
     campoasiento(i - 1) = datcheques.Recordset.Fields("idasiento")
     camponroasiento(i - 1) = datcheques.Recordset.Fields("nroasiento")
     campoperiodo(i - 1) = datcheques.Recordset.Fields("perinicial")
    
    
     If campo1 = "" Or IsNull(campo1) = True Then
            campo = "__/__/____"
     Else
            campo = Str(campo1)
     End If
     
     If campo1a = "" Or IsNull(campo1a) = True Then
            campoa = "__/__/____"
     Else
            campoa = Str(campo1a)
     End If
     
     If campo1b = "" Or IsNull(campo1b) = True Then
            campob = "__/__/____"
     Else
            campob = Str(campo1b)
     End If
     
     campo2 = campo2 + "_______________"
     campo2 = Left(campo2, 15)
    
     campo3 = Mid(campo3, 1, 60)
     
     
     campo4 = Format(campo4, "#0.00")
     campo4 = "__________" + campo4
     campo4 = Right(campo4, 13)
     campo5 = Format(campo5, "#0.00")
     campo5 = "__________" + campo5
     campo5 = Right(campo5, 13)
       
       
     lista = campo & "   " & campoa & "   " & campob & "   " & campo2 & "    " & campo4 & "    " & campo5 & "      " & campo3 & "---" & campo6
     List1.AddItem (lista)
     
Rem     If IsNull(datcheques.Recordset.Fields("conciliado")) = True Then List1.Selected(i) = False
     
     datcheques.Recordset.MoveNext

Loop

For x = i + 1 To 99999
    tipo(x) = ""
    idinstru(x) = 0
    cancelado(x) = False
    campoasiento(x) = 0
    camponroasiento(x) = 0
    as_cuenta(x) = 0
    as_debe(x) = 0
    as_haber(0) = 0
    as_desc(x) = ""
    as_cc(x) = 0
Next x

For x = 0 To List1.ListCount - 1
        List1.Selected(x) = cancelado(x)
Next x
    
End Sub

Private Sub nuevo_Click()

    datconciliacion.RecordSource = "select conciliacion.* from conciliacion where empresa = " & login.empresaact & "  "
    datconciliacion.Refresh
    datconciliacion.Recordset.AddNew
    fechaconc.Value = Date
    DTPicker1.Value = Date
    Text2.SetFocus
    datconciliacion.Recordset.Fields("empresa") = login.empresaact
    datconciliacion.Recordset.Fields("cerrada") = "N"
    datconciliacion.Recordset.UpdateBatch adAffectCurrent
    text1.Text = datconciliacion.Recordset.Fields("nro")

End Sub

Private Sub Option1_Click()

Rem    Call llena_Click

End Sub

Private Sub Option2_Click()

 Rem   Call llena_Click

End Sub

Private Sub Option3_Click()

 Rem   Call llena_Click

End Sub

Private Sub Option4_Click()

 Rem   Call llena_Click

End Sub

Private Sub Option5_Click()

 Rem   Call llena_Click
    
End Sub

Private Sub Option6_Click()


Rem    Call llena_Click
    
End Sub



Private Sub Option7_Click()
 Rem Call llena_Click
End Sub

Private Sub Option8_Click()
 Rem Call llena_Click
End Sub

Private Sub Option9_Click()

 Rem  Call llena_Click

End Sub
