VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form frmproveedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proveedores Maestro"
   ClientHeight    =   6465
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   10275
   Icon            =   "frmproveedores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6465
   ScaleWidth      =   10275
   Begin VB.CommandButton verificacuenta 
      Caption         =   "verificacuenta"
      Height          =   255
      Left            =   1200
      TabIndex        =   47
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmproveedores.frx":12BFE
      Height          =   495
      Left            =   120
      TabIndex        =   46
      Top             =   5280
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.Frame Frame5 
      Height          =   1695
      Left            =   6360
      TabIndex        =   45
      Top             =   4560
      Width           =   3615
      Begin KewlButtonz.KewlButtons aceptar 
         Height          =   615
         Left            =   120
         TabIndex        =   48
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
         MICON           =   "frmproveedores.frx":12C18
         PICN            =   "frmproveedores.frx":12C34
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
         Left            =   1200
         TabIndex        =   49
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
         MICON           =   "frmproveedores.frx":13646
         PICN            =   "frmproveedores.frx":13662
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons elminar 
         Height          =   615
         Left            =   2280
         TabIndex        =   50
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
         MICON           =   "frmproveedores.frx":14074
         PICN            =   "frmproveedores.frx":14090
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
         Left            =   720
         TabIndex        =   53
         Top             =   960
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
         MICON           =   "frmproveedores.frx":17482
         PICN            =   "frmproveedores.frx":1749E
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
         Left            =   1800
         TabIndex        =   54
         Top             =   960
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
         MICON           =   "frmproveedores.frx":17EB0
         PICN            =   "frmproveedores.frx":17ECC
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
   Begin VB.Frame Frame4 
      Caption         =   "Otros"
      Height          =   1695
      Left            =   120
      TabIndex        =   41
      Top             =   4560
      Width           =   6135
      Begin vbskpro.Skinner Skinner1 
         Left            =   120
         Top             =   1440
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
         LcK2            =   $"frmproveedores.frx":18A16
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
      Begin MSComCtl2.DTPicker fechacai 
         Height          =   255
         Left            =   2280
         TabIndex        =   38
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   161873921
         CurrentDate     =   39356
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   17
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   37
         Top             =   360
         Width           =   3255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Fecha Venc.CAI:"
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
         Left            =   600
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Último CAI:"
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
         Left            =   840
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   15
         Left            =   2280
         TabIndex        =   39
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Plazo Venc.Facturas:"
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
         Left            =   120
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Impuestos"
      Height          =   1575
      Left            =   4440
      TabIndex        =   31
      Top             =   3000
      Width           =   4575
      Begin VB.CommandButton Command1 
         Caption         =   "Tabla Retención Ganancias:"
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
         Left            =   120
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   14
         Left            =   3120
         TabIndex        =   36
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Tabla Retención IVA:"
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
         Left            =   840
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Tabla Retención IIBB:"
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
         Left            =   840
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   13
         Left            =   3120
         TabIndex        =   33
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   12
         Left            =   3120
         TabIndex        =   32
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Contabilidad"
      Height          =   1215
      Left            =   120
      TabIndex        =   26
      Top             =   3000
      Width           =   4215
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   11
         Left            =   2280
         TabIndex        =   30
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   2280
         TabIndex        =   29
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cuenta de Gasto  :"
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
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cuenta Proveedor:"
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
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.CommandButton llena 
      Caption         =   "llena"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   4200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   7560
         TabIndex        =   24
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   7560
         MaxLength       =   13
         TabIndex        =   23
         Top             =   1320
         Width           =   1575
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmproveedores.frx":18A25
         Height          =   315
         Left            =   7560
         TabIndex        =   22
         Top             =   240
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
         Caption         =   "Código:"
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
         Left            =   6240
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   19
         Top             =   2400
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   18
         Top             =   2040
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   17
         Top             =   1680
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   16
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   15
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   14
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   13
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
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1680
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
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1320
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
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Identificación Impositiva"
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
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   600
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
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         TabIndex        =   2
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
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSAdodcLib.Adodc datcondtrib 
      Height          =   330
      Left            =   120
      Top             =   4920
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
      Left            =   120
      Top             =   3600
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
            Picture         =   "frmproveedores.frx":18A3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmproveedores.frx":19451
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc datcuentas 
      Height          =   330
      Left            =   1800
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
   Begin KewlButtonz.KewlButtons siguiente 
      Height          =   495
      Left            =   9240
      TabIndex        =   51
      TabStop         =   0   'False
      ToolTipText     =   "Siguiente"
      Top             =   3480
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
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
      MICON           =   "frmproveedores.frx":19E63
      PICN            =   "frmproveedores.frx":19E7F
      PICH            =   "frmproveedores.frx":206E1
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
      Height          =   495
      Left            =   9240
      TabIndex        =   52
      TabStop         =   0   'False
      ToolTipText     =   "Anterior"
      Top             =   3960
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
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
      MICON           =   "frmproveedores.frx":26F43
      PICN            =   "frmproveedores.frx":26F5F
      PICH            =   "frmproveedores.frx":2D7C1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Crystal.CrystalReport CrystalReporte 
      Left            =   1320
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowTitle     =   "Libro IVA Compras"
   End
   Begin MSRDC.MSRDC reporte 
      Height          =   375
      Left            =   4560
      Top             =   120
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
   Begin MSAdodcLib.Adodc datverifica 
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
End
Attribute VB_Name = "frmproveedores"
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
       
        If text1(0).Text = "" Then
            MsgBox "No puede ingresar un valor Nulo en la RAZONSOCIAL", vbCritical, "Error"
            text1(0).SetFocus
            Exit Sub
        End If
        
        If flagnuevo = 0 Then
                bases.datbasemenu.Recordset.AddNew
                bases.datbasemenu.Recordset.Fields("empresa") = login.empresaact
        End If
        bases.datbasemenu.Recordset.Fields("razonsocial") = text1(0).Text
        bases.datbasemenu.Recordset.Fields("domicilio") = text1(1).Text
        bases.datbasemenu.Recordset.Fields("localidad") = text1(2).Text
        bases.datbasemenu.Recordset.Fields("codpostal") = text1(3).Text
        bases.datbasemenu.Recordset.Fields("telefono") = text1(4).Text
        bases.datbasemenu.Recordset.Fields("email") = text1(5).Text
        bases.datbasemenu.Recordset.Fields("contacto") = text1(6).Text
        If DataCombo1.Text = "" Then
            MsgBox "Ingrese Condición Tributaria", vbCritical, "Error"
            DataCombo1.SetFocus
            bases.datbasemenu.Recordset.Delete adAffectCurrent
            Exit Sub
        End If
        bases.datbasemenu.Recordset.Fields("tipoiva") = DataCombo1.BoundText
        bases.datbasemenu.Recordset.Fields("cuit") = text1(7).Text
Rem        If flagnuevo = 1 Then
Rem              bases.datbasemenu.Recordset.Fields("codproveedor") = Text1(9).Text
Rem        End If
        If text1(10).Text = "" Then text1(10).Text = 0
        bases.datbasemenu.Recordset.Fields("codcontable") = text1(10).Text
        If text1(11).Text = "" Then text1(11).Text = 0
        bases.datbasemenu.Recordset.Fields("codcontablegastos") = text1(11).Text
        bases.datbasemenu.Recordset.Fields("fechavenccai") = fechacai.Value
        If text1(15).Text = "" Then text1(15).Text = 0
        bases.datbasemenu.Recordset.Fields("plazovencfacturas") = text1(15).Text
        bases.datbasemenu.Recordset.Fields("ultimocai") = text1(17).Text
        bases.datbasemenu.Recordset.Fields("zonafiscal") = text1(8).Text
        bases.datbasemenu.Recordset.Fields("tablaretiva") = text1(12).Text
        bases.datbasemenu.Recordset.Fields("tablaretib") = text1(13).Text
        bases.datbasemenu.Recordset.Fields("tablaretgan") = text1(14).Text

        bases.datbasemenu.Recordset.UpdateBatch adAffectCurrent
        MsgBox "Almacenado Correctamente", vbInformation, "Guardar"
        lista_proveedores.lista = ""
        flagnuevo = 0
        Call llena_Click
        text1(0).SetFocus

End Sub

Private Sub anterior_Click()
On Error Resume Next
    If bases.datbasemenu.Recordset.RecordCount = 1 Then
        bases.datbasemenu.RecordSource = "select proveedores.* from proveedores where empresa = " & empresareal & " order by razonsocial"
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
        text1(0).SetFocus
        
End Sub

Private Sub cerrar_Click()

Unload Me

End Sub

Private Sub Command1_GotFocus(Index As Integer)

    text1(0).SetFocus

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
        text1(7).SetFocus
    End If

End Sub


Private Sub elminar_Click()

    datverifica.RecordSource = "SELECT cuit, proveedor, empresa From librocompras where empresa = " & login.empresaact & " GROUP BY cuit, proveedor, empresa"
    datverifica.Refresh
    datverifica.Recordset.Filter = "proveedor = '" & text1(0).Text & "' and cuit = '" & text1(7).Text & "'"
    If datverifica.Recordset.EOF = False Then
        MsgBox "Este Proveedor tiene Comprobantes imputados, no se puede eliminar", vbCritical, "Error"
        Exit Sub
    End If
    
    datverifica.RecordSource = "SELECT codproveedor, nomproveedor, empresa From ordendepagoabonan where empresa = " & login.empresaact & " GROUP BY codproveedor, nomproveedor, empresa"
    datverifica.Refresh
    datverifica.Recordset.Filter = "nomproveedor = '" & text1(0).Text & "' and codproveedor = " & text1(9).Text & ""
    If datverifica.Recordset.EOF = False Then
        MsgBox "Este Proveedor tiene Ordenes imputados, no se puede eliminar", vbCritical, "Error"
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
        text1(15).SetFocus
    End If
End Sub

Private Sub Form_Load()
Aplicar_skin Me

frmproveedores.Top = 0
frmproveedores.Left = 0

If login.provaltas = "N" Or login.provmodi = "N" Then
    aceptar.Enabled = False
Else
    aceptar.Enabled = True
End If

If login.provbajas = "N" Then
    elminar.Enabled = False
Else
    elminar.Enabled = True
End If

ventana.menu = 0
navega = 0

datcondtrib.ConnectionString = login.conexiontotal
datcuentas.ConnectionString = login.conexiontotal
datverifica.ConnectionString = login.conexiontotal
datcondtrib.RecordSource = "select condtrib.* from condtrib"
datcondtrib.Refresh

If login.empresaact1 > 0 Then
    empresareal = login.empresaact1
Else
    empresareal = login.empresaact
End If

bases.datbasemenu.ConnectionString = login.conexiontotal
bases.datbasemenu1.ConnectionString = login.conexiontotal
bases.datbasemenu.RecordSource = "select proveedores.* from proveedores where empresa = " & empresareal & " order by razonsocial"
bases.datbasemenu.Refresh

flagnuevo = 0

End Sub

Private Sub menaccion_Click()

End Sub

Private Sub KewlButtons1_Click()
Dim tabla As String
Dim ruta As String

ruta = "\Empresa" + Right(Str(login.empresaact), Len(Str(login.empresaact)) - 1)

    mensa = MsgBox("Ordena por Codigo(Si), Razon Social(No)", vbYesNoCancel, "Listar")
    If mensa = vbYes Then reporte.SQL = "SELECT proveedores.codproveedor, proveedores.razonsocial, proveedores.tipoiva, proveedores.cuit, proveedores.domicilio, proveedores.localidad, EMPRESA.razonsocial FROM { oj contablesql.dbo.proveedores proveedores INNER JOIN contablesql.dbo.EMPRESA EMPRESA ON proveedores.empresa = EMPRESA.empresa} where proveedores.empresa = " & empresareal & " ORDER BY proveedores.codproveedor ASC"
    If mensa = vbNo Then reporte.SQL = "SELECT proveedores.codproveedor, proveedores.razonsocial, proveedores.tipoiva, proveedores.cuit, proveedores.domicilio, proveedores.localidad, EMPRESA.razonsocial FROM { oj contablesql.dbo.proveedores proveedores INNER JOIN contablesql.dbo.EMPRESA EMPRESA ON proveedores.empresa = EMPRESA.empresa} where proveedores.empresa = " & empresareal & " ORDER BY proveedores.razonsocial ASC"
    If mensa = vbCancel Then Exit Sub

Rem reporte.SQL = "SELECT clientes.codcliente, clientes.razonsocial, clientes.tipoiva, clientes.cuit, clientes.domicilio, clientes.localidad, EMPRESA.razonsocial FROM { oj contablesql.dbo.clientes clientes INNER JOIN contablesql.dbo.EMPRESA EMPRESA ON clientes.empresa = EMPRESA.empresa} where clientes.empresa = " & login.empresaact & " ORDER BY clientes.codcliente ASC"
tabla = reporte.SQL

With CrystalReporte
    .ReportFileName = App.Path & ruta + "\proveedores.rpt"
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

Private Sub llena_Click()
On Error Resume Next
        For X = 0 To 11
            text1(X).Text = ""
        Next X
        
   If navega = 0 Then
        bases.datbasemenu.RecordSource = "select proveedores.* from proveedores where empresa = " & empresareal & " and razonsocial = '" & lista_proveedores.lista & "' "
        bases.datbasemenu.Refresh
        text1(0).Text = lista_proveedores.lista
   Else
        text1(0).Text = bases.datbasemenu.Recordset.Fields("razonsocial")
   End If
        If IsNull(bases.datbasemenu.Recordset.Fields("domicilio")) Then
                text1(1).Text = ""
        Else
                text1(1).Text = bases.datbasemenu.Recordset.Fields("domicilio")
        End If
        If IsNull(bases.datbasemenu.Recordset.Fields("localidad")) Then
                text1(2).Text = ""
        Else
                text1(2).Text = bases.datbasemenu.Recordset.Fields("localidad")
        End If
        If IsNull(bases.datbasemenu.Recordset.Fields("codpostal")) Then
                text1(3).Text = ""
        Else
                text1(3).Text = bases.datbasemenu.Recordset.Fields("codpostal")
        End If
        If IsNull(bases.datbasemenu.Recordset.Fields("telefono")) Then
                text1(4).Text = ""
        Else
                text1(4).Text = bases.datbasemenu.Recordset.Fields("telefono")
        End If
        If IsNull(bases.datbasemenu.Recordset.Fields("email")) Then
                text1(5).Text = ""
        Else
                text1(5).Text = bases.datbasemenu.Recordset.Fields("email")
        End If
        If IsNull(bases.datbasemenu.Recordset.Fields("contacto")) Then
                text1(6).Text = ""
        Else
                text1(6).Text = bases.datbasemenu.Recordset.Fields("contacto")
        End If
        DataCombo1.BoundText = bases.datbasemenu.Recordset.Fields("tipoiva")
        If IsNull(bases.datbasemenu.Recordset.Fields("cuit")) Then
                text1(7).Text = ""
        Else
                text1(7).Text = bases.datbasemenu.Recordset.Fields("cuit")
        End If
        text1(9).Text = bases.datbasemenu.Recordset.Fields("codproveedor")
        If IsNull(bases.datbasemenu.Recordset.Fields("codcontable")) Then
                text1(10).Text = ""
        Else
                text1(10).Text = bases.datbasemenu.Recordset.Fields("codcontable")
        End If
        If IsNull(bases.datbasemenu.Recordset.Fields("codcontablegastos")) Then
                text1(11).Text = ""
        Else
                text1(11).Text = bases.datbasemenu.Recordset.Fields("codcontablegastos")
        End If
        If IsNull(bases.datbasemenu.Recordset.Fields("fechavenccai")) = True Then
             fechacai.Value = #12/31/2100#
        Else
            fechacai.Value = bases.datbasemenu.Recordset.Fields("fechavenccai")
        End If
        If IsNull(bases.datbasemenu.Recordset.Fields("plazovencfacturas")) Then
                text1(15).Text = ""
        Else
                text1(15).Text = bases.datbasemenu.Recordset.Fields("plazovencfacturas")
        End If
        If IsNull(bases.datbasemenu.Recordset.Fields("ultimocai")) Then
                text1(17).Text = ""
        Else
                text1(17).Text = bases.datbasemenu.Recordset.Fields("ultimocai")
        End If
        If IsNull(bases.datbasemenu.Recordset.Fields("zonafiscal")) Then
                text1(8).Text = ""
        Else
                text1(8).Text = bases.datbasemenu.Recordset.Fields("zonafiscal")
        End If
        If IsNull(bases.datbasemenu.Recordset.Fields("tablaretiva")) Then
                text1(12).Text = ""
        Else
                text1(12).Text = bases.datbasemenu.Recordset.Fields("tablaretiva")
        End If
        If IsNull(bases.datbasemenu.Recordset.Fields("tablaretib")) Then
                text1(13).Text = ""
        Else
                text1(13).Text = bases.datbasemenu.Recordset.Fields("tablaretib")
        End If
        If IsNull(bases.datbasemenu.Recordset.Fields("tablaretgan")) Then
                text1(14).Text = ""
        Else
                text1(14).Text = bases.datbasemenu.Recordset.Fields("tablaretgan")
        End If
        
        ventana.menu = 0
        navega = 0
        
End Sub

Private Sub siguiente_Click()
On Error Resume Next
    
    If bases.datbasemenu.Recordset.RecordCount = 1 Then
        bases.datbasemenu.RecordSource = "select proveedores.* from proveedores where empresa = " & empresareal & " order by razonsocial"
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

    If Index = 7 Then
        If Len(text1(Index).Text) = 2 Or Len(text1(Index).Text) = 11 Then
            text1(Index).Text = text1(Index).Text + "-"
            SendKeys "{end}", False
        End If
    End If

End Sub

Private Sub Text1_GotFocus(Index As Integer)

    If ventana.menu = 1 And Index = 0 Then
        Call llena_Click
    End If
    

    If ventana.menu = 5 And Index = 10 Then
        ventana.menu = 0
        text1(10).Text = lista_cuentas.cuentacont
    End If
    If ventana.menu = 5 And Index = 11 Then
        ventana.menu = 0
        text1(11).Text = lista_cuentas.cuentacont
    End If
    

    
    
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

    If Index = 7 And KeyAscii = 27 Then
            salir = 1
            Call Cancelar_Click
    End If



    If KeyAscii = 13 Then
        KeyAscii = 0
        If Index = 10 Or Index = 11 Then
            posicion = Index
            Call verificacuenta_Click
        End If
        
        salir = 0
        SendKeys "{tab}", False
    End If


End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 38 Then
        text1(Index - 1).SetFocus
        Exit Sub
    End If
    
    If KeyCode = 114 And Index = 0 Then
        ventana.menu = 1
        lista_proveedores.Show
    End If
    
    If KeyCode = 117 And Index = 0 Then
        Call siguiente_Click
    End If
    If KeyCode = 116 And Index = 0 Then
        Call anterior_Click
    End If
    
    If KeyCode = 114 And Index = 10 Then
        lista_cuentas.cuentacont = text1(10).Text
        ventana.menu = 5
        lista_cuentas.Show
    End If
    If KeyCode = 114 And Index = 11 Then
        lista_cuentas.cuentacont = text1(11).Text
        ventana.menu = 5
        lista_cuentas.Show
    End If

End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim invalido As Integer
    If Index = 7 And salir = 0 And DataGrid1.Columns(2).Text = "1" Then
            mensa = verifica_cuit(text1(7).Text, invalido)
 Rem           If invalido = 1 Then
 Rem               Text1(7).Text = ""
 Rem               Text1(7).SetFocus
 Rem           End If
            posicion = bases.datbasemenu.Recordset.AbsolutePosition
            bases.datbasemenu1.RecordSource = "select cuit,empresa,razonsocial from proveedores where empresa = " & empresareal & " and cuit = '" & text1(7).Text & "'"
            bases.datbasemenu1.Refresh
            If bases.datbasemenu1.Recordset.EOF = False Then
                If bases.datbasemenu1.Recordset.Fields("razonsocial") <> text1(0).Text Then
                    mensa = MsgBox("Ya existe otro proveedor con el mismo Nro. de CUIT", vbCritical, "Error")
                    text1(7).SetFocus
                End If
            End If
    End If
End Sub

Private Sub verificacuenta_Click()

    If text1(posicion) = "" Then Exit Sub

    datcuentas.ConnectionString = login.conexiontotal
    datcuentas.RecordSource = "select cuentas.* from cuentas"
    datcuentas.Refresh
    datcuentas.Recordset.Filter = "empre = " & login.empresaact & " and imp = 'S' and [cod contable] = " & text1(posicion).Text & " and inicioper = '" & login.iper & "'"
    
    If datcuentas.Recordset.EOF = True Then
        MsgBox "No Existe esta cuenta contable", vbCritical, "Verificar"
        text1(posicion).Text = ""
        text1(posicion).SetFocus
    End If
    
End Sub
