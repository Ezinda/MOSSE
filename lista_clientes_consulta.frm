VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form lista_clientes_consulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Clientes Consulta"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14670
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   14670
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "Anotaciones de Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   9240
      TabIndex        =   27
      Top             =   600
      Visible         =   0   'False
      Width           =   9735
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   29
         Text            =   "lista_clientes_consulta.frx":0000
         Top             =   600
         Width           =   9375
      End
      Begin VB.CommandButton Command9 
         Caption         =   "X"
         Height          =   375
         Left            =   9120
         TabIndex        =   28
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Valores Recibidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   2160
      TabIndex        =   19
      Top             =   1080
      Visible         =   0   'False
      Width           =   11535
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "lista_clientes_consulta.frx":0006
         Height          =   4095
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   7223
         _Version        =   393216
         AllowUpdate     =   0   'False
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
      Begin VB.CommandButton Command5 
         Caption         =   "X"
         Height          =   375
         Left            =   11040
         TabIndex        =   20
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cuenta Corriente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   4440
      TabIndex        =   6
      Top             =   6120
      Width           =   11775
      Begin VB.CheckBox Check2 
         Caption         =   "Envio por Mail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8760
         TabIndex        =   30
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Provincia:"
         Height          =   375
         Left            =   360
         TabIndex        =   24
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "-->"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10920
         TabIndex        =   23
         Top             =   120
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Filtra Incobrables"
         Height          =   375
         Left            =   10080
         TabIndex        =   22
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Totales"
         Height          =   255
         Left            =   8760
         TabIndex        =   17
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Detallado"
         Height          =   255
         Left            =   8760
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
      Begin KewlButtonz.KewlButtons KewlButtons1 
         Height          =   615
         Left            =   2760
         TabIndex        =   13
         Top             =   1680
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "Cta.Cte. con Aplicaciones, No requiere filtro de Fecha"
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
         MICON           =   "lista_clientes_consulta.frx":001F
         PICN            =   "lista_clientes_consulta.frx":003B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Hasta Fecha:"
         Height          =   375
         Left            =   3360
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Desde Fecha:"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DesdeFecha 
         Height          =   375
         Left            =   1680
         TabIndex        =   9
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   49676289
         CurrentDate     =   42198
      End
      Begin MSComCtl2.DTPicker HastaFecha 
         Height          =   375
         Left            =   4680
         TabIndex        =   10
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   49676289
         CurrentDate     =   42198
      End
      Begin KewlButtonz.KewlButtons historicoctacte 
         Height          =   615
         Left            =   360
         TabIndex        =   11
         Top             =   1680
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "Detalle Cta.Cte."
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
         MICON           =   "lista_clientes_consulta.frx":048D
         PICN            =   "lista_clientes_consulta.frx":04A9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Crystal.CrystalReport CrystalReporte 
         Left            =   1440
         Top             =   1200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Presupusto de Venta"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrinterCollation=   0
         PrintFileLinesPerPage=   60
         WindowShowGroupTree=   -1  'True
         WindowShowSearchBtn=   -1  'True
      End
      Begin MSRDC.MSRDC reporte 
         Height          =   375
         Left            =   120
         Top             =   1200
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
      Begin MSAdodcLib.Adodc datfiltro 
         Height          =   330
         Left            =   2040
         Top             =   1320
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
         Caption         =   "datitemsnv"
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
      Begin KewlButtonz.KewlButtons KewlButtons2 
         Height          =   615
         Left            =   6480
         TabIndex        =   14
         Top             =   1680
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "Compr.Pendientes de Saldar"
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
         MICON           =   "lista_clientes_consulta.frx":08FB
         PICN            =   "lista_clientes_consulta.frx":0917
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons KewlButtons3 
         Height          =   615
         Left            =   6480
         TabIndex        =   15
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "Saldo hasta Fecha"
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
         MICON           =   "lista_clientes_consulta.frx":0D69
         PICN            =   "lista_clientes_consulta.frx":0D85
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons KewlButtons4 
         Height          =   615
         Left            =   9000
         TabIndex        =   18
         Top             =   1680
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "Valores Recibidos"
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
         MICON           =   "lista_clientes_consulta.frx":11D7
         PICN            =   "lista_clientes_consulta.frx":11F3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSAdodcLib.Adodc datvalores 
         Height          =   330
         Left            =   2280
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
         Caption         =   "datitemsnv"
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
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "lista_clientes_consulta.frx":1064C
         Height          =   420
         Left            =   1680
         TabIndex        =   25
         Top             =   1080
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   741
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "nombre"
         BoundColumn     =   "nombre"
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
      Begin MSAdodcLib.Adodc datprovincia 
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
         Caption         =   "datitemsnv"
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
   Begin VB.CommandButton salir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   13200
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid datagrid1 
      Bindings        =   "lista_clientes_consulta.frx":10667
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   9551
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
      HeadLines       =   1
      RowHeight       =   15
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   14670
      _ExtentX        =   25876
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      Begin VB.CommandButton Command8 
         Caption         =   "&Historial de Venta"
         Height          =   375
         Left            =   9480
         TabIndex        =   26
         Top             =   120
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   8160
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Exporta a Excel"
         Height          =   375
         Left            =   11280
         TabIndex        =   5
         Top             =   120
         Width           =   1815
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
         Height          =   405
         Left            =   1080
         TabIndex        =   1
         Top             =   120
         Width           =   6975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar:"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   120
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc datcliente 
      Height          =   330
      Left            =   0
      Top             =   4440
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
   Begin MSAdodcLib.Adodc datparametros 
      Height          =   450
      Left            =   0
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   794
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
Attribute VB_Name = "lista_clientes_consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cuentacont As String
Dim cuenta(99999) As Integer


' -----------------------------------------------------------------------------------------
' \\ --  Descripción       : Exportar DataGrid a Excel
' \\ --  Controles         : Un Datagrid, un CommandButton y la referencia a ADO
' \\ --  Autor             : Luciano Lodola -- http://www.recursosvisualbasic.com.ar/
' -----------------------------------------------------------------------------------------

' -- Variables para la base de datos
Dim cnn         As Connection
Dim rs          As Recordset
' -- Variables para Excel
Dim Obj_Excel   As Object
Dim Obj_Libro   As Object
Dim Obj_Hoja    As Object


Private Sub Combo1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And Combo1.Text = "Sin Limite" Then
        KeyAscii = 0
             xquery = "SELECT ALIAS_0.ID, ALIAS_0.CODIGO, ALIAS_3.NOMBRE AS RAZONSOCIAL, ALIAS_3.CUIT, ALIAS_6.CALLE, ISNULL(ALIAS_6.CALLE, '') + '-' + ISNULL(V_CIUDAD_.NOMBRE, '') " & _
              "+ '-' + ISNULL(V_PROVINCIA_.NOMBRE, '') AS DOMICILIO, ALIAS_0.DENOMINACION, ALIAS_5.NOMBRE AS ZONA, ALIAS_7.NUMERO AS TELEFONO, " & _
              "ALIAS_8.DIRECCIONELECTRONICA AS MAIL, V_TIPOPAGO_.NOMBRE AS TP, V_TIPOPAGO_.OBSERVACION AS TipoPago, " & _
              "V_TIPOPAGO_.ACTIVESTATUS AS ACTIVESTATUSTP, V_PROVINCIA_.NOMBRE AS Provincia, V_CIUDAD_.NOMBRE AS Ciudad, " & _
              "ALIAS_0.CODIGO + ' ' + ALIAS_3.CUIT + ' ' + ALIAS_3.NOMBRE AS CONCATENADO, V_EZI_POCICION_IVA_CLIENTES.CODIGO AS IVA, V_TIPOPAGO_.ID AS IDPAGO, " & _
              "v_ezi_pos_impuestos.COEFICIENTE AS Alic_IIBB, v_ezi_pos_impuestos.NOMBRE AS Cond_IIBB, v_ezi_pos_impuestos.EXENCION AS Exencion_IIBB " & _
              "FROM         V_TIPOPAGO_ RIGHT OUTER JOIN " & _
              "V_CLIENTE AS ALIAS_0 WITH (nolock) LEFT OUTER JOIN " & _
              "v_ezi_pos_impuestos ON ALIAS_0.ID = v_ezi_pos_impuestos.idcliente LEFT OUTER JOIN " & _
              "V_EZI_POCICION_IVA_CLIENTES ON ALIAS_0.ID = V_EZI_POCICION_IVA_CLIENTES.idcliente ON V_TIPOPAGO_.ID = ALIAS_0.TIPOPAGO_ID LEFT OUTER JOIN " & _
              "V_PERSONA AS ALIAS_3 WITH (nolock) ON ALIAS_0.ENTEASOCIADO_ID = ALIAS_3.ID LEFT OUTER JOIN " & _
              "V_CIUDAD_ RIGHT OUTER JOIN " & _
              "V_DOMICILIO_ AS ALIAS_6 ON V_CIUDAD_.ID = ALIAS_6.CIUDAD_ID LEFT OUTER JOIN " & _
              "V_PROVINCIA_ ON ALIAS_6.PROVINCIA_ID = V_PROVINCIA_.ID ON ALIAS_0.DOMICILIOFACTURACION_ID = ALIAS_6.ID LEFT OUTER JOIN " & _
              "V_TELEFONO_ AS ALIAS_7 ON ALIAS_3.TELEFONOPRINCIPAL_ID = ALIAS_7.ID LEFT OUTER JOIN " & _
              "V_DIRECCIONELECTRONICA_ AS ALIAS_8 ON ALIAS_3.DIRECELECTRONICAPRINCIPAL_ID = ALIAS_8.ID LEFT OUTER JOIN " & _
              "V_ZONA_ AS ALIAS_5 ON ALIAS_0.ZONA_ID = ALIAS_5.ID " & _
              "ORDER BY RAZONSOCIAL"

        datcliente.RecordSource = xquery
        datcliente.Refresh
        
                    datagrid1.Columns(0).Visible = False
            datagrid1.Columns(7).Visible = False
            'DataGrid1.Columns(9).Visible = False
            datagrid1.Columns(10).Visible = False
            datagrid1.Columns(12).Visible = False
            datagrid1.Columns(15).Visible = False
            datagrid1.Columns(17).Visible = False
            datagrid1.Columns(18).Visible = False
            datagrid1.Columns(1).Width = 1000
            datagrid1.Columns(2).Width = 3500


        
    End If


End Sub

' -------------------------------------------------------------------------------
' \\ -- Botón para Ejecutar la función que exporta los datos del datagrid a excel
' -------------------------------------------------------------------------------
Private Sub Command3_Click()
 '   On Error Resume Next
    
    Dim i   As Integer
    Dim j   As Integer
    
    n_Filas = datcliente.Recordset.RecordCount
    ' -- Colocar el cursor de espera mientras se exportan los datos
    Me.MousePointer = vbHourglass
    
    If n_Filas = 0 Then
        MsgBox "No hay datos para exportar a excel. Se ha indicado 0 en el parámetro Filas ": Exit Sub
    Else
        
   '     Set o_Excel = CreateObject("Excel.Application")
    'Set o_Libro = o_Excel.Workbooks.Add
    'Set o_Hoja = o_Libro.Worksheets.Add
        
        ' -- Crear nueva instancia de Excel
        Set Obj_Excel = CreateObject("Excel.Application")
        ' -- Agregar nuevo libro}
        'xruta = App.Path + "\Clientes.xls"
        'Set Obj_Libro = Obj_Excel.Workbooks.Open(App.Path)
    
        ' -- Referencia a la Hoja activa ( la que añade por defecto Excel )
        Set o_Libro = Obj_Excel.Workbooks.Add
        Set o_Hoja = o_Libro.Worksheets.Add
        Set Obj_Hoja = Obj_Excel.ActiveSheet
   
        iCol = 0
        ' --  Recorrer el Datagrid ( Las columnas )
        For i = 0 To datagrid1.Columns.Count - 1
          If i = 0 Or i = 7 Or i = 10 Or i = 12 Or i = 15 Or i = 17 Then GoTo sigue
            If datagrid1.Columns(i).Visible Then
                ' -- Incrementar índice de columna
                iCol = iCol + 1
                ' -- Obtener el caption de la columna
                Obj_Hoja.Cells(1, iCol) = datagrid1.Columns(i).Caption
                ' -- Recorrer las filas
                For j = 0 To n_Filas - 1
                    ' -- Asignar el valor a la celda del Excel
                    Obj_Hoja.Cells(j + 2, iCol) = _
                    datagrid1.Columns(i).CellValue(datagrid1.GetBookmark(j))
                Next
            End If
sigue:
        Next
        
        ' -- Hacer excel visible
        Obj_Excel.Visible = True
        
        ' -- Opcional : colocar en negrita y de color rojo los enbezados en la hoja
        With Obj_Hoja
            .Rows(1).Font.Bold = True
            .Rows(1).Font.Color = vbRed
            ' -- Autoajustar las cabeceras
            .Columns("A:Z").AutoFit
        End With
    End If

    ' -- Eliminar las variables de objeto excel
    Set Obj_Hoja = Nothing
    Set Obj_Libro = Nothing
    Set Obj_Excel = Nothing
    
    ' -- Restaurar cursor
    Me.MousePointer = vbDefault

End Sub

  


Private Sub Command5_Click()

    Frame2.Visible = False

End Sub

Private Sub Command6_Click()

    If Frame1.Width <> Command6.Width Then
        Command6.Caption = "<--"
        Command6.Left = 0
        Frame1.Width = Command6.Width
        Frame1.Height = Command6.Height + 100
    Else
        Command6.Caption = "-->"
        Command6.Left = 10920
        Frame1.Width = 11775
        Frame1.Height = 2415
    End If
    

End Sub

Private Sub Command8_Click()
On Error Resume Next
    menu = 1
      query = "SELECT     R.claveprimaria, R.clienteid, R.cliente, R.fechadelcomprobante, RD.referenciaproducto, RD.nombre_producto, RD.cantidadoriginal, RD.unidaddemedida, R.presupuestobase , NVD.preciou " & _
              "FROM         ud_ezi_puntodeventa_encabezado AS R WITH (nolock) INNER JOIN  " & _
              "ud_ezi_puntodeventa_detalle_rem AS RD WITH (nolock) ON R.claveprimaria = RD.claveprimaria INNER JOIN " & _
              "ud_ezi_puntodeventa_encabezado AS NV WITH (nolock) ON R.presupuestobase = NV.claveprimaria INNER JOIN " & _
              "ud_ezi_puntodeventa_detalle_notav AS NVD WITH (nolock) ON RD.idproducto = NVD.idproducto AND RD.item = NVD.item AND NV.id = NVD.claveprimaria " & _
              "WHERE     (R.numeradorinterno = 'Remito de Venta') and R.clienteid =  '" & datagrid1.Columns("id").Text & "'" & _
              "ORDER BY R.fechadelcomprobante DESC  "
    lista_historial.Show

End Sub

Private Sub Command9_Click()

Frame3.Visible = False

End Sub

Private Sub DataGrid1_Click()

'Frame3.Visible = False


End Sub

Private Sub DataGrid1_DblClick()

On Error Resume Next

    Frame3.Visible = True
    Frame3.Caption = datagrid1.Columns("razonsocial").Text
    Text2.Text = datagrid1.Columns("anotaciones").Text
    

End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

Frame3.Visible = False

End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)

Frame3.Visible = False

End Sub

Private Sub Form_Load()
Aplicar_skin Me

MiFuncionDeAjuste Me, True

yventana = Inicio.Height / 2 - 1000
xventana = Inicio.Width / 2

'lista_clientes_consulta.Top = yventana - lista_clientes_consulta.Height / 2
'lista_clientes_consulta.Left = xventana - lista_clientes_consulta.Width / 2

Combo1.AddItem ("50")
Combo1.AddItem ("Sin Limite")
Option1.Value = True

Combo1.Text = "50"
Check1.Value = 0
Check2.Value = 0

DesdeFecha.Value = Date - Day(Date) + 1
HastaFecha.Value = Date


datprovincia.ConnectionString = login.conexiontotal
datprovincia.RecordSource = "select ' TODAS' as Nombre Union All SELECT     NOMBRE from V_PROVINCIA_ " & _
                            "WHERE     (BO_PLACE_ID = '{DA4D078D-0ACA-4B3E-BD72-2EB9F4EE145C}') and NOMBRE <> 'No Usar' order by NOMBRE"
datprovincia.Refresh
datprovincia.Recordset.MoveFirst
DataCombo3.BoundText = datprovincia.Recordset.Fields("nombre")

   datparametros.ConnectionString = login.conexiontotal
   datparametros.RecordSource = "select * from ud_ezi_parametros_pos with (nolock) where sucursal = '" & login.nomsucursal & "' "
   datparametros.Refresh
   
   
   xcodclientefiltra = datparametros.Recordset.Fields("codclientefiltra")
   If xcodclientefiltra = "01" Then xcontrocliente = "88447B8E-14FE-4D60-9622-B22F6C735701"  ' tucuman
   If xcodclientefiltra = "04" Then xcontrocliente = "4234CA46-B2BE-4690-AC6A-F0DE206F94A9"  ' salta
   If xcodclientefiltra = "03" Then xcontrocliente = "AEC7FBAC-63F7-4404-9512-033D0961D9BC"  ' jujuy


datcliente.ConnectionString = login.conexiontotal

             
     xquery = "SELECT     TOP (50) ALIAS_0.ID, ALIAS_0.CODIGO, ALIAS_3.NOMBRE AS RAZONSOCIAL, ALIAS_3.CUIT, ALIAS_6.CALLE, ISNULL(ALIAS_6.CALLE, '') " & _
              "+ '-' + ISNULL(V_CIUDAD.NOMBRE, '') + '-' + ISNULL(V_PROVINCIA.NOMBRE, '') AS DOMICILIO, ALIAS_0.DENOMINACION, ALIAS_5.NOMBRE AS ZONA, " & _
              "ALIAS_7.NUMERO AS TELEFONO, ALIAS_8.DIRECCIONELECTRONICA AS MAIL, V_TIPOPAGO.NOMBRE AS TP, V_TIPOPAGO.OBSERVACION AS TipoPago, " & _
              "V_TIPOPAGO.ACTIVESTATUS AS ACTIVESTATUSTP, V_PROVINCIA.NOMBRE AS Provincia, V_CIUDAD.NOMBRE AS Ciudad, " & _
              "ALIAS_0.CODIGO + ' ' + ALIAS_3.CUIT + ' ' + ALIAS_3.NOMBRE AS CONCATENADO, V_EZI_POCICION_IVA_CLIENTES.CODIGO AS IVA, V_TIPOPAGO.ID AS IDPAGO, " & _
              "v_ezi_pos_impuestos.COEFICIENTE AS Alic_IIBB, v_ezi_pos_impuestos.NOMBRE AS Cond_IIBB, v_ezi_pos_impuestos.EXENCION AS Exencion_IIBB, V_UD_CLIENTE.Anotaciones " & _
              "FROM         V_EZI_POCICION_IVA_CLIENTES RIGHT OUTER JOIN " & _
              "V_CLIENTE AS ALIAS_0 WITH (nolock) LEFT OUTER JOIN " & _
              "V_UD_CLIENTE with (nolock) ON ALIAS_0.BOEXTENSION_ID = V_UD_CLIENTE.ID LEFT OUTER JOIN " & _
              "v_ezi_pos_impuestos ON ALIAS_0.ID = v_ezi_pos_impuestos.idcliente ON V_EZI_POCICION_IVA_CLIENTES.idcliente = ALIAS_0.ID LEFT OUTER JOIN " & _
              "V_TIPOPAGO with (nolock) ON ALIAS_0.TIPOPAGO_ID = V_TIPOPAGO.ID LEFT OUTER JOIN " & _
              "V_PERSONA AS ALIAS_3 WITH (nolock) ON ALIAS_0.ENTEASOCIADO_ID = ALIAS_3.ID LEFT OUTER JOIN " & _
              "V_CIUDAD with (nolock) RIGHT OUTER JOIN " & _
              "V_DOMICILIO  AS ALIAS_6 with (nolock) ON V_CIUDAD.ID = ALIAS_6.CIUDAD_ID LEFT OUTER JOIN " & _
              "V_PROVINCIA with (nolock) ON ALIAS_6.PROVINCIA_ID = V_PROVINCIA.ID ON ALIAS_0.DOMICILIOFACTURACION_ID = ALIAS_6.ID LEFT OUTER JOIN " & _
              "V_TELEFONO  AS ALIAS_7 with (nolock) ON ALIAS_3.TELEFONOPRINCIPAL_ID = ALIAS_7.ID LEFT OUTER JOIN " & _
              "V_DIRECCIONELECTRONICA  AS ALIAS_8  with (nolock) ON ALIAS_3.DIRECELECTRONICAPRINCIPAL_ID = ALIAS_8.ID LEFT OUTER JOIN " & _
              "V_ZONA  AS ALIAS_5 with (nolock) ON ALIAS_0.ZONA_ID = ALIAS_5.ID " & _
              "ORDER BY RAZONSOCIAL"
              

datcliente.RecordSource = xquery
datcliente.Refresh


            datagrid1.Columns(0).Visible = False
            datagrid1.Columns(7).Visible = False
            'DataGrid1.Columns(9).Visible = False
            datagrid1.Columns(4).Visible = False
            datagrid1.Columns(10).Visible = False
            datagrid1.Columns(12).Visible = False
            datagrid1.Columns(15).Visible = False
            datagrid1.Columns(17).Visible = False
            datagrid1.Columns(18).Visible = False
            datagrid1.Columns(1).Width = 1000
            datagrid1.Columns(2).Width = 3500
            datagrid1.Columns(6).Width = 3500

 
End Sub

Private Sub Form_Resize()
On Error Resume Next
    datagrid1.Width = lista_clientes_consulta.Width - 200
    datagrid1.Height = lista_clientes_consulta.Height - 1000

End Sub

Private Sub historicoctacte_Click()
On Error Resume Next
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim ruta As String


    datfiltro.ConnectionString = login.conexiontotal
    datfiltro.RecordSource = "select * from ud_ezi_pos_filtroec"
    datfiltro.Refresh
    
    datfiltro.Recordset.Fields("desdefecha") = Replace(Str(Year(DesdeFecha.Value)), " ", "") + Right("0" + Replace(Str(Month(DesdeFecha.Value)), " ", ""), 2) + Right("0" + Replace(Str(Day(DesdeFecha.Value)), " ", ""), 2) + "000000000"
    datfiltro.Recordset.Fields("hastafecha") = Replace(Str(Year(HastaFecha.Value)), " ", "") + Right("0" + Replace(Str(Month(HastaFecha.Value)), " ", ""), 2) + Right("0" + Replace(Str(Day(HastaFecha.Value)), " ", ""), 2) + "000000000"
    If login.nomsucursal = "EMPORIO" Then datfiltro.Recordset.Fields("empresa") = "EL EMP.TUCUMAN"
    If login.nomsucursal = "EMPORIOZIP" Then datfiltro.Recordset.Fields("empresa") = "COMPRADOR Tucuman"
    If login.nomsucursal = "TUCUMAN" Then datfiltro.Recordset.Fields("empresa") = "MOM"
    If login.nomsucursal = "TUCUMANZIP" Then datfiltro.Recordset.Fields("empresa") = "DIM TOLEDO VAL"

If Text1.Text = "" Then
    datfiltro.Recordset.Fields("cliente") = ""

   
Else
    datfiltro.Recordset.Fields("cliente") = datagrid1.Columns("codigo").Text

End If

datfiltro.Recordset.UpdateBatch adAffectCurrent
If Check1.Value = 0 Then
  If DataCombo3.Text = " TODAS" Then
        reporte.SQL = "SELECT v_ezi_pos_ctacte.TIPO, v_ezi_pos_ctacte.CODIGO, v_ezi_pos_ctacte.NOMBREDESTINATARIO, v_ezi_pos_ctacte.NUMERODOCUMENTO, v_ezi_pos_ctacte.NOMBRE, v_ezi_pos_ctacte.FECHAACTUAL, v_ezi_pos_ctacte.TOTAL, v_ezi_pos_ctacte.NOMCLASIFICADOR FROM MMOSSE.dbo.v_ezi_pos_ctacte v_ezi_pos_ctacte  ORDER BY v_ezi_pos_ctacte.CODIGO ASC, v_ezi_pos_ctacte.FECHAACTUAL ASC"
  Else
        reporte.SQL = "SELECT v_ezi_pos_ctacte.TIPO, v_ezi_pos_ctacte.CODIGO, v_ezi_pos_ctacte.NOMBREDESTINATARIO, v_ezi_pos_ctacte.NUMERODOCUMENTO, v_ezi_pos_ctacte.NOMBRE, v_ezi_pos_ctacte.FECHAACTUAL, v_ezi_pos_ctacte.TOTAL, v_ezi_pos_ctacte.NOMCLASIFICADOR FROM MMOSSE.dbo.v_ezi_pos_ctacte v_ezi_pos_ctacte WHERE v_ezi_pos_ctacte.PROVINCIA = '" & DataCombo3.Text & "' ORDER BY v_ezi_pos_ctacte.CODIGO ASC, v_ezi_pos_ctacte.FECHAACTUAL ASC"
  End If
Else
  If DataCombo3.Text = " TODAS" Then
        reporte.SQL = "SELECT v_ezi_pos_ctacte.TIPO, v_ezi_pos_ctacte.CODIGO, v_ezi_pos_ctacte.NOMBREDESTINATARIO, v_ezi_pos_ctacte.NUMERODOCUMENTO, v_ezi_pos_ctacte.NOMBRE, v_ezi_pos_ctacte.FECHAACTUAL, v_ezi_pos_ctacte.TOTAL, v_ezi_pos_ctacte.NOMCLASIFICADOR FROM MMOSSE.dbo.v_ezi_pos_ctacte v_ezi_pos_ctacte where v_ezi_pos_ctacte.incobrables = 'F'  ORDER BY v_ezi_pos_ctacte.CODIGO ASC, v_ezi_pos_ctacte.FECHAACTUAL ASC"
  Else
        reporte.SQL = "SELECT v_ezi_pos_ctacte.TIPO, v_ezi_pos_ctacte.CODIGO, v_ezi_pos_ctacte.NOMBREDESTINATARIO, v_ezi_pos_ctacte.NUMERODOCUMENTO, v_ezi_pos_ctacte.NOMBRE, v_ezi_pos_ctacte.FECHAACTUAL, v_ezi_pos_ctacte.TOTAL, v_ezi_pos_ctacte.NOMCLASIFICADOR FROM MMOSSE.dbo.v_ezi_pos_ctacte v_ezi_pos_ctacte where v_ezi_pos_ctacte.incobrables = 'F'  AND v_ezi_pos_ctacte.PROVINCIA = '" & DataCombo3.Text & "' ORDER BY v_ezi_pos_ctacte.CODIGO ASC, v_ezi_pos_ctacte.FECHAACTUAL ASC"
  End If
End If

tabla = reporte.SQL

With CrystalReporte
    .PrinterCollation = crptCollated
    If Option1.Value = True Then
        .ReportFileName = App.Path & "\Reporte_Historico_ctacte.rpt"
    Else
        .ReportFileName = App.Path & "\Reporte_Historico_ctacte_totales.rpt"
    End If
    .WindowTitle = "Historico Cta.Cte. Clientes"
    .Formulas(0) = "desdefecha=""" & DesdeFecha.Value & """"
    .Formulas(1) = "hastafecha=""" & HastaFecha.Value & """"
    .Formulas(2) = "provincia=""" & DataCombo3.Text & """"
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
  If Check2.Value = 1 Then
     Kill ("c:\util\*.pdf")
     Kill ("c:\util\*.rtf")


    .Destination = crptToFile
    .PrintFileType = crptRTF

    .PrintFileName = "c:\util\CtaCte " + Replace(datagrid1.Columns(2), "*", "") + ".pdf"

    .WindowState = crptNormal
    .Action = 1
    
'    PDFCreator_CreatePDF Me, "c:\util\CtaCte " + Replace(datagrid1.Columns(2), "*", "") + ".rtf", "c:\util\CtaCte " + Replace(datagrid1.Columns(2), "*", "") + ".pdf"
  Else
    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .Action = 1
  End If

    
End With

If Check2.Value = 1 Then
 Set oApp = New Outlook.Application
 Set myItem = oApp.CreateItem(Outlook.OlItemType.olMailItem)
 Set myAttachments = myItem.Attachments
 myAttachments.Add "c:\util\ctacte " + Replace(datagrid1.Columns(2), "*", "") + ".pdf", 4, 2, " "
 myItem.Display

 Set oApp = Nothing
 Set myItem = Nothing
 Set myAttachments = Nothing
End If

    
Exit Sub

fuera:
    
    MsgBox "Reporte no Encontado, o error de configuracion de reporte", vbCritical, "Error"





End Sub

Private Sub KewlButtons1_Click()
On Error Resume Next
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim ruta As String


    datfiltro.ConnectionString = login.conexiontotal
    datfiltro.RecordSource = "select * from ud_ezi_pos_filtroec"
    datfiltro.Refresh
    
    datfiltro.Recordset.Fields("desdefecha") = Replace(Str(Year(DesdeFecha.Value)), " ", "") + Right("0" + Replace(Str(Month(DesdeFecha.Value)), " ", ""), 2) + Right("0" + Replace(Str(Day(DesdeFecha.Value)), " ", ""), 2) + "000000000"
    datfiltro.Recordset.Fields("hastafecha") = Replace(Str(Year(HastaFecha.Value)), " ", "") + Right("0" + Replace(Str(Month(HastaFecha.Value)), " ", ""), 2) + Right("0" + Replace(Str(Day(HastaFecha.Value)), " ", ""), 2) + "000000000"
    If login.nomsucursal = "EMPORIO" Then datfiltro.Recordset.Fields("empresa") = "EL EMP.TUCUMAN"
    If login.nomsucursal = "EMPORIOZIP" Then datfiltro.Recordset.Fields("empresa") = "COMPRADOR Tucuman"
    If login.nomsucursal = "TUCUMAN" Then datfiltro.Recordset.Fields("empresa") = ""
    If login.nomsucursal = "TUCUMANZIP" Then datfiltro.Recordset.Fields("empresa") = "DIM TOLEDO VAL"
        
    



If Text1.Text = "" Then
    datfiltro.Recordset.Fields("cliente") = ""

   
Else
    datfiltro.Recordset.Fields("cliente") = datagrid1.Columns("codigo").Text

End If

datfiltro.Recordset.UpdateBatch adAffectCurrent

If Check1.Value = 0 Then
 If DataCombo3.Text = " TODAS" Then
    reporte.SQL = "SELECT v_ezi_pos_ctacte_conimputaciones.COD_CLIENTE, v_ezi_pos_ctacte_conimputaciones.NOM_CLIENTE, v_ezi_pos_ctacte_conimputaciones.FECHAEMISION, v_ezi_pos_ctacte_conimputaciones.FECHAVENCIMIENTO, v_ezi_pos_ctacte_conimputaciones.IMPORTE_TOTAL, v_ezi_pos_ctacte_conimputaciones.SALDO, v_ezi_pos_ctacte_conimputaciones.NOMCLASIFICADOR, v_ezi_pos_ctacte_conimputaciones.FECHAEMISION_IMP, v_ezi_pos_ctacte_conimputaciones.IMPORTE, v_ezi_pos_ctacte_conimputaciones.COMPROBANTE, v_ezi_pos_ctacte_conimputaciones.grupo FROM MMOSSE.dbo.v_ezi_pos_ctacte_conimputaciones v_ezi_pos_ctacte_conimputaciones ORDER BY v_ezi_pos_ctacte_conimputaciones.NOM_CLIENTE ASC, v_ezi_pos_ctacte_conimputaciones.grupo ASC, v_ezi_pos_ctacte_conimputaciones.FECHAEMISION ASC"
 Else
     reporte.SQL = "SELECT v_ezi_pos_ctacte_conimputaciones.COD_CLIENTE, v_ezi_pos_ctacte_conimputaciones.NOM_CLIENTE, v_ezi_pos_ctacte_conimputaciones.FECHAEMISION, v_ezi_pos_ctacte_conimputaciones.FECHAVENCIMIENTO, v_ezi_pos_ctacte_conimputaciones.IMPORTE_TOTAL, v_ezi_pos_ctacte_conimputaciones.SALDO, v_ezi_pos_ctacte_conimputaciones.NOMCLASIFICADOR, v_ezi_pos_ctacte_conimputaciones.FECHAEMISION_IMP, v_ezi_pos_ctacte_conimputaciones.IMPORTE, v_ezi_pos_ctacte_conimputaciones.COMPROBANTE, v_ezi_pos_ctacte_conimputaciones.grupo FROM MMOSSE.dbo.v_ezi_pos_ctacte_conimputaciones v_ezi_pos_ctacte_conimputaciones WHERE v_ezi_pos_ctacte_conimputaciones.PROVINCIA = '" & DataCombo3.Text & "' ORDER BY v_ezi_pos_ctacte_conimputaciones.NOM_CLIENTE ASC, v_ezi_pos_ctacte_conimputaciones.grupo ASC, v_ezi_pos_ctacte_conimputaciones.FECHAEMISION ASC"
 End If
Else
 If DataCombo3.Text = " TODAS" Then
    reporte.SQL = "SELECT v_ezi_pos_ctacte_conimputaciones.COD_CLIENTE, v_ezi_pos_ctacte_conimputaciones.NOM_CLIENTE, v_ezi_pos_ctacte_conimputaciones.FECHAEMISION, v_ezi_pos_ctacte_conimputaciones.FECHAVENCIMIENTO, v_ezi_pos_ctacte_conimputaciones.IMPORTE_TOTAL, v_ezi_pos_ctacte_conimputaciones.SALDO, v_ezi_pos_ctacte_conimputaciones.NOMCLASIFICADOR, v_ezi_pos_ctacte_conimputaciones.FECHAEMISION_IMP, v_ezi_pos_ctacte_conimputaciones.IMPORTE, v_ezi_pos_ctacte_conimputaciones.COMPROBANTE, v_ezi_pos_ctacte_conimputaciones.grupo FROM MMOSSE.dbo.v_ezi_pos_ctacte_conimputaciones v_ezi_pos_ctacte_conimputaciones where v_ezi_pos_ctacte_conimputaciones.incobrables = 'F' ORDER BY v_ezi_pos_ctacte_conimputaciones.NOM_CLIENTE ASC, v_ezi_pos_ctacte_conimputaciones.grupo ASC, v_ezi_pos_ctacte_conimputaciones.FECHAEMISION ASC"
 Else
    reporte.SQL = "SELECT v_ezi_pos_ctacte_conimputaciones.COD_CLIENTE, v_ezi_pos_ctacte_conimputaciones.NOM_CLIENTE, v_ezi_pos_ctacte_conimputaciones.FECHAEMISION, v_ezi_pos_ctacte_conimputaciones.FECHAVENCIMIENTO, v_ezi_pos_ctacte_conimputaciones.IMPORTE_TOTAL, v_ezi_pos_ctacte_conimputaciones.SALDO, v_ezi_pos_ctacte_conimputaciones.NOMCLASIFICADOR, v_ezi_pos_ctacte_conimputaciones.FECHAEMISION_IMP, v_ezi_pos_ctacte_conimputaciones.IMPORTE, v_ezi_pos_ctacte_conimputaciones.COMPROBANTE, v_ezi_pos_ctacte_conimputaciones.grupo FROM MMOSSE.dbo.v_ezi_pos_ctacte_conimputaciones v_ezi_pos_ctacte_conimputaciones where v_ezi_pos_ctacte_conimputaciones.incobrables = 'F' AND v_ezi_pos_ctacte_conimputaciones.PROVINCIA = '" & DataCombo3.Text & "' ORDER BY v_ezi_pos_ctacte_conimputaciones.NOM_CLIENTE ASC, v_ezi_pos_ctacte_conimputaciones.grupo ASC, v_ezi_pos_ctacte_conimputaciones.FECHAEMISION ASC"
 End If
End If


tabla = reporte.SQL

With CrystalReporte
    .PrinterCollation = crptCollated
    .ReportFileName = App.Path & "\Reporte_ctacte_conimputaciones.rpt"
    .WindowTitle = "Historico Cta.Cte. Clientes"
'    .Formulas(0) = "desdefecha=""" & DesdeFecha.Value & """"
'    .Formulas(1) = "hastafecha=""" & HastaFecha.Value & """"
    .Formulas(2) = "provincia=""" & DataCombo3.Text & """"
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
  If Check2.Value = 1 Then
     Kill ("c:\util\*.pdf")
     Kill ("c:\util\*.rtf")


    .Destination = crptToFile
    .PrintFileType = crptRTF

    .PrintFileName = "c:\util\CtaCte " + Replace(datagrid1.Columns(2), "*", "") + ".pdf"

    .WindowState = crptNormal
    .Action = 1
    
'    PDFCreator_CreatePDF Me, "c:\util\CtaCte " + Replace(datagrid1.Columns(2), "*", "") + ".rtf", "c:\util\CtaCte " + Replace(datagrid1.Columns(2), "*", "") + ".pdf"
  Else
    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .Action = 1
  End If

    
End With

If Check2.Value = 1 Then
 Set oApp = New Outlook.Application
 Set myItem = oApp.CreateItem(Outlook.OlItemType.olMailItem)
 Set myAttachments = myItem.Attachments
 myAttachments.Add "c:\util\ctacte " + Replace(datagrid1.Columns(2), "*", "") + ".pdf", 4, 2, " "
 myItem.Display

 Set oApp = Nothing
 Set myItem = Nothing
 Set myAttachments = Nothing
End If

    
Exit Sub

fuera:
    
    MsgBox "Reporte no Encontado, o error de configuracion de reporte", vbCritical, "Error"






End Sub

Private Sub KewlButtons2_Click()
On Error Resume Next
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim ruta As String


    datfiltro.ConnectionString = login.conexiontotal
    datfiltro.RecordSource = "select * from ud_ezi_pos_filtroec"
    datfiltro.Refresh
    
    xdesdefecha = Replace(Str(Year(DesdeFecha.Value)), " ", "") + Right("0" + Replace(Str(Month(DesdeFecha.Value)), " ", ""), 2) + Right("0" + Replace(Str(Day(DesdeFecha.Value)), " ", ""), 2)
    xhastafecha = Replace(Str(Year(HastaFecha.Value)), " ", "") + Right("0" + Replace(Str(Month(HastaFecha.Value)), " ", ""), 2) + Right("0" + Replace(Str(Day(HastaFecha.Value)), " ", ""), 2)
    If login.nomsucursal = "EMPORIO" Then xcp = "EL EMP.TUCUMAN"
    If login.nomsucursal = "EMPORIOZIP" Then xcp = "COMPRADOR Tucuman"
    If login.nomsucursal = "TUCUMAN" Then xcp = ""
    If login.nomsucursal = "TUCUMANZIP" Then xcp = "DIM TOLEDO VAL"
        
    xCliente = datagrid1.Columns(0).Text
If Check1.Value = 0 Then
    reporte.SQL = "SELECT v_ezi_pos_ctacte_pendientes.TIPO, v_ezi_pos_ctacte_pendientes.DESCRIPCION, v_ezi_pos_ctacte_pendientes.NOM_CLIENTE, v_ezi_pos_ctacte_pendientes.FECHAEMISION, v_ezi_pos_ctacte_pendientes.SALDO FROM MMOSSE.dbo.v_ezi_pos_ctacte_pendientes v_ezi_pos_ctacte_pendientes " & _
                  "where v_ezi_pos_ctacte_pendientes.id = '" & xCliente & "' and substring(v_ezi_pos_ctacte_pendientes.FECHAEMISION,1,8) >= '" & xdesdefecha & "' and substring(v_ezi_pos_ctacte_pendientes.FECHAEMISION,1,8) <= '" & xhastafecha & "' " & _
                  "ORDER BY v_ezi_pos_ctacte_pendientes.NOM_CLIENTE ASC, v_ezi_pos_ctacte_pendientes.FECHAEMISION asc"
Else
    reporte.SQL = "SELECT v_ezi_pos_ctacte_pendientes.TIPO, v_ezi_pos_ctacte_pendientes.DESCRIPCION, v_ezi_pos_ctacte_pendientes.NOM_CLIENTE, v_ezi_pos_ctacte_pendientes.FECHAEMISION, v_ezi_pos_ctacte_pendientes.SALDO FROM MMOSSE.dbo.v_ezi_pos_ctacte_pendientes v_ezi_pos_ctacte_pendientes " & _
                  "where v_ezi_pos_ctacte_pendientes.id = '" & xCliente & "' and substring(v_ezi_pos_ctacte_pendientes.FECHAEMISION,1,8) >= '" & xdesdefecha & "' and substring(v_ezi_pos_ctacte_pendientes.FECHAEMISION,1,8) <= '" & xhastafecha & "' " & _
                  "and v_ezi_pos_ctacte_pendientes.incobrables = 'F' " & _
                  "ORDER BY v_ezi_pos_ctacte_pendientes.NOM_CLIENTE ASC, v_ezi_pos_ctacte_pendientes.FECHAEMISION asc"
End If

'reporte.SQL = "SELECT v_ezi_pos_ctacte.TIPO, v_ezi_pos_ctacte.CODIGO, v_ezi_pos_ctacte.NOMBREDESTINATARIO, v_ezi_pos_ctacte.NUMERODOCUMENTO, v_ezi_pos_ctacte.NOMBRE, v_ezi_pos_ctacte.FECHAACTUAL, v_ezi_pos_ctacte.TOTAL, v_ezi_pos_ctacte.NOMCLASIFICADOR FROM MMOSSE.dbo.v_ezi_pos_ctacte v_ezi_pos_ctacte ORDER BY v_ezi_pos_ctacte.CODIGO ASC, v_ezi_pos_ctacte.FECHAACTUAL ASC"

tabla = reporte.SQL

With CrystalReporte
    .PrinterCollation = crptCollated
    .ReportFileName = App.Path & "\Reporte_ctacte_pendientes.rpt"
    .WindowTitle = "Comprobantes Pendientes de Saldar"
    .Formulas(0) = "desdefecha=""" & DesdeFecha.Value & """"
    .Formulas(1) = "hastafecha=""" & HastaFecha.Value & """"
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
  If Check2.Value = 1 Then
     Kill ("c:\util\*.pdf")
     Kill ("c:\util\*.rtf")


    .Destination = crptToFile
    .PrintFileType = crptRTF

    .PrintFileName = "c:\util\CtaCte " + Replace(datagrid1.Columns(2), "*", "") + ".pdf"

    .WindowState = crptNormal
    .Action = 1
    
'    PDFCreator_CreatePDF Me, "c:\util\CtaCte " + Replace(datagrid1.Columns(2), "*", "") + ".rtf", "c:\util\CtaCte " + Replace(datagrid1.Columns(2), "*", "") + ".pdf"
  Else
    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .Action = 1
  End If

    
End With

If Check2.Value = 1 Then
 Set oApp = New Outlook.Application
 Set myItem = oApp.CreateItem(Outlook.OlItemType.olMailItem)
 Set myAttachments = myItem.Attachments
 myAttachments.Add "c:\util\ctacte " + Replace(datagrid1.Columns(2), "*", "") + ".pdf", 4, 2, " "
 myItem.Display

 Set oApp = Nothing
 Set myItem = Nothing
 Set myAttachments = Nothing
End If

    
Exit Sub

fuera:
    
    MsgBox "Reporte no Encontado, o error de configuracion de reporte", vbCritical, "Error"





End Sub

Private Sub KewlButtons3_Click()
On Error Resume Next
Dim crApp As New CRAXDRT.Application
Dim crReport As New CRAXDRT.Report
Dim tabla As String
Dim ruta As String


    datfiltro.ConnectionString = login.conexiontotal
    datfiltro.RecordSource = "select * from ud_ezi_pos_filtroec"
    datfiltro.Refresh
    
    datfiltro.Recordset.Fields("desdefecha") = Replace(Str(Year(DesdeFecha.Value)), " ", "") + Right("0" + Replace(Str(Month(DesdeFecha.Value)), " ", ""), 2) + Right("0" + Replace(Str(Day(DesdeFecha.Value)), " ", ""), 2) + "000000000"
    datfiltro.Recordset.Fields("hastafecha") = Replace(Str(Year(HastaFecha.Value + 1)), " ", "") + Right("0" + Replace(Str(Month(HastaFecha.Value + 1)), " ", ""), 2) + Right("0" + Replace(Str(Day(HastaFecha.Value + 1)), " ", ""), 2) + "000000000"
    If login.nomsucursal = "EMPORIO" Then datfiltro.Recordset.Fields("empresa") = "EL EMP.TUCUMAN"
    If login.nomsucursal = "EMPORIOZIP" Then datfiltro.Recordset.Fields("empresa") = "COMPRADOR Tucuman"
    If login.nomsucursal = "TUCUMAN" Then datfiltro.Recordset.Fields("empresa") = ""
    If login.nomsucursal = "TUCUMANZIP" Then datfiltro.Recordset.Fields("empresa") = "DIM TOLEDO VAL"
        

If Text1.Text = "" Then
    datfiltro.Recordset.Fields("cliente") = ""
Else
    datfiltro.Recordset.Fields("cliente") = datagrid1.Columns("codigo").Text
End If

If Option1.Value = False Then
    datfiltro.Recordset.Fields("cliente") = ""
End If

datfiltro.Recordset.UpdateBatch adAffectCurrent
If Check1.Value = 0 Then
  If DataCombo3.Text = " TODAS" Then
        reporte.SQL = "SELECT v_ezi_pos_ctacte_conimputaciones_afecha.COD_CLIENTE, v_ezi_pos_ctacte_conimputaciones_afecha.NOM_CLIENTE, v_ezi_pos_ctacte_conimputaciones_afecha.FECHAEMISION, v_ezi_pos_ctacte_conimputaciones_afecha.FECHAVENCIMIENTO, v_ezi_pos_ctacte_conimputaciones_afecha.IMPORTE_TOTAL, v_ezi_pos_ctacte_conimputaciones_afecha.SALDO, v_ezi_pos_ctacte_conimputaciones_afecha.NOMCLASIFICADOR, v_ezi_pos_ctacte_conimputaciones_afecha.FECHAEMISION_IMP, v_ezi_pos_ctacte_conimputaciones_afecha.IMPORTE, v_ezi_pos_ctacte_conimputaciones_afecha.COMPROBANTE, v_ezi_pos_ctacte_conimputaciones_afecha.grupo FROM MMOSSE.dbo.v_ezi_pos_ctacte_conimputaciones_afecha v_ezi_pos_ctacte_conimputaciones_afecha ORDER BY v_ezi_pos_ctacte_conimputaciones_afecha.NOM_CLIENTE ASC, v_ezi_pos_ctacte_conimputaciones_afecha.COD_CLIENTE ASC, v_ezi_pos_ctacte_conimputaciones_afecha.grupo ASC, v_ezi_pos_ctacte_conimputaciones_afecha.FECHAEMISION ASC"
  Else
        reporte.SQL = "SELECT v_ezi_pos_ctacte_conimputaciones_afecha.COD_CLIENTE, v_ezi_pos_ctacte_conimputaciones_afecha.NOM_CLIENTE, v_ezi_pos_ctacte_conimputaciones_afecha.FECHAEMISION, v_ezi_pos_ctacte_conimputaciones_afecha.FECHAVENCIMIENTO, v_ezi_pos_ctacte_conimputaciones_afecha.IMPORTE_TOTAL, v_ezi_pos_ctacte_conimputaciones_afecha.SALDO, v_ezi_pos_ctacte_conimputaciones_afecha.NOMCLASIFICADOR, v_ezi_pos_ctacte_conimputaciones_afecha.FECHAEMISION_IMP, v_ezi_pos_ctacte_conimputaciones_afecha.IMPORTE, v_ezi_pos_ctacte_conimputaciones_afecha.COMPROBANTE, v_ezi_pos_ctacte_conimputaciones_afecha.grupo FROM MMOSSE.dbo.v_ezi_pos_ctacte_conimputaciones_afecha v_ezi_pos_ctacte_conimputaciones_afecha where v_ezi_pos_ctacte_conimputaciones_afecha.PROVINCIA = '" & DataCombo3.Text & "' ORDER BY v_ezi_pos_ctacte_conimputaciones_afecha.NOM_CLIENTE ASC, v_ezi_pos_ctacte_conimputaciones_afecha.COD_CLIENTE ASC " & _
        ", v_ezi_pos_ctacte_conimputaciones_afecha.grupo ASC, v_ezi_pos_ctacte_conimputaciones_afecha.FECHAEMISION ASC"
  End If
Else
  If DataCombo3.Text = " TODAS" Then
        reporte.SQL = "SELECT v_ezi_pos_ctacte_conimputaciones_afecha.COD_CLIENTE, v_ezi_pos_ctacte_conimputaciones_afecha.NOM_CLIENTE, v_ezi_pos_ctacte_conimputaciones_afecha.FECHAEMISION, v_ezi_pos_ctacte_conimputaciones_afecha.FECHAVENCIMIENTO, v_ezi_pos_ctacte_conimputaciones_afecha.IMPORTE_TOTAL, v_ezi_pos_ctacte_conimputaciones_afecha.SALDO, v_ezi_pos_ctacte_conimputaciones_afecha.NOMCLASIFICADOR, v_ezi_pos_ctacte_conimputaciones_afecha.FECHAEMISION_IMP, v_ezi_pos_ctacte_conimputaciones_afecha.IMPORTE, v_ezi_pos_ctacte_conimputaciones_afecha.COMPROBANTE, v_ezi_pos_ctacte_conimputaciones_afecha.grupo FROM MMOSSE.dbo.v_ezi_pos_ctacte_conimputaciones_afecha v_ezi_pos_ctacte_conimputaciones_afecha where v_ezi_pos_ctacte_conimputaciones_afecha.incobrables = 'F' ORDER BY v_ezi_pos_ctacte_conimputaciones_afecha.NOM_CLIENTE ASC, v_ezi_pos_ctacte_conimputaciones_afecha.COD_CLIENTE ASC, v_ezi_pos_ctacte_conimputaciones_afecha.grupo ASC, v_ezi_pos_ctacte_conimputaciones_afecha.FECHAEMISION ASC"
  Else
        reporte.SQL = "SELECT v_ezi_pos_ctacte_conimputaciones_afecha.COD_CLIENTE, v_ezi_pos_ctacte_conimputaciones_afecha.NOM_CLIENTE, v_ezi_pos_ctacte_conimputaciones_afecha.FECHAEMISION, v_ezi_pos_ctacte_conimputaciones_afecha.FECHAVENCIMIENTO, v_ezi_pos_ctacte_conimputaciones_afecha.IMPORTE_TOTAL, v_ezi_pos_ctacte_conimputaciones_afecha.SALDO, v_ezi_pos_ctacte_conimputaciones_afecha.NOMCLASIFICADOR, v_ezi_pos_ctacte_conimputaciones_afecha.FECHAEMISION_IMP, v_ezi_pos_ctacte_conimputaciones_afecha.IMPORTE, v_ezi_pos_ctacte_conimputaciones_afecha.COMPROBANTE, v_ezi_pos_ctacte_conimputaciones_afecha.grupo FROM MMOSSE.dbo.v_ezi_pos_ctacte_conimputaciones_afecha v_ezi_pos_ctacte_conimputaciones_afecha where v_ezi_pos_ctacte_conimputaciones_afecha.incobrables = 'F' and  v_ezi_pos_ctacte_conimputaciones_afecha.PROVINCIA = '" & DataCombo3.Text & "' " & _
                      "ORDER BY v_ezi_pos_ctacte_conimputaciones_afecha.NOM_CLIENTE ASC, v_ezi_pos_ctacte_conimputaciones_afecha.COD_CLIENTE ASC, v_ezi_pos_ctacte_conimputaciones_afecha.grupo ASC, v_ezi_pos_ctacte_conimputaciones_afecha.FECHAEMISION ASC"
  End If
End If

tabla = reporte.SQL

With CrystalReporte
    .PrinterCollation = crptCollated
    If Option1.Value = True Then
        .ReportFileName = App.Path & "\Reporte_ctacte_conimputaciones_afecha.rpt"
    Else
        .ReportFileName = App.Path & "\Reporte_ctacte_conimputaciones_afecha_totales.rpt"
    End If
    .WindowTitle = "Historico Cta.Cte. Clientes a fecha"
'    .Formulas(0) = "desdefecha=""" & DesdeFecha.Value & """"
    .Formulas(1) = "hastafecha=""" & HastaFecha.Value & """"
    .Formulas(2) = "provincia=""" & DataCombo3.Text & """"
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla

  If Check2.Value = 1 Then
     Kill ("c:\util\*.pdf")
     Kill ("c:\util\*.rtf")


    .Destination = crptToFile
    .PrintFileType = crptRTF

    .PrintFileName = "c:\util\CtaCte " + Replace(datagrid1.Columns(2), "*", "") + ".pdf"

    .WindowState = crptNormal
    .Action = 1
    
'    PDFCreator_CreatePDF Me, "c:\util\CtaCte " + Replace(datagrid1.Columns(2), "*", "") + ".rtf", "c:\util\CtaCte " + Replace(datagrid1.Columns(2), "*", "") + ".pdf"
  Else
    .Destination = crptToWindow
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .Action = 1
  End If

    
End With

If Check2.Value = 1 Then
 Set oApp = New Outlook.Application
 Set myItem = oApp.CreateItem(Outlook.OlItemType.olMailItem)
 Set myAttachments = myItem.Attachments
 myAttachments.Add "c:\util\ctacte " + Replace(datagrid1.Columns(2), "*", "") + ".pdf", 4, 2, " "
 myItem.Display

 Set oApp = Nothing
 Set myItem = Nothing
 Set myAttachments = Nothing
End If
 
  
 
Exit Sub

fuera:
    
    MsgBox "Reporte no Encontado, o error de configuracion de reporte", vbCritical, "Error"


End Sub

Private Sub KewlButtons4_Click()
On Error Resume Next

Frame2.Visible = True

Frame2.Caption = "Valores recibidos de : " + datagrid1.Columns(1).Text + " - " + datagrid1.Columns(2).Text
datvalores.ConnectionString = login.conexiontotal

datvalores.RecordSource = "SELECT     ALIAS_0.DATOS, right(ALIAS_0.FECHAEMISION,2) + '/' +SUBSTRING(ALIAS_0.FECHAEMISION,5,2) + '/' + LEFT(ALIAS_0.FECHAEMISION,4) as Fechaemision, RIGHT(ALIAS_0.FECHAVEN, 2) + '/' + SUBSTRING(ALIAS_0.FECHAVEN, 5, 2) + '/' + LEFT(ALIAS_0.FECHAVEN, 4) AS FECHAVEN, " & _
                          "ALIAS_0.VALOR2_IMPORTE AS IMPORTE, CASE WHEN pasado = 'T' THEN 'No' ELSE 'Si' END AS En_Cartera,ALIAS_1.TIPOVALOR,  V_CLIENTE_.ID, " & _
                          "V_PERSONA_.NOMBRE AS Cliente " & _
                          "FROM         V_PERSONA_ RIGHT OUTER JOIN " & _
                          "V_CLIENTE_ ON V_PERSONA_.ID = V_CLIENTE_.ENTEASOCIADO_ID RIGHT OUTER JOIN  " & _
                          "V_ITEMVALOR_ AS ALIAS_0 ON V_CLIENTE_.ID = ALIAS_0.OPERADORCOMERCIAL_ID LEFT OUTER JOIN " & _
                          "V_TIPOVALOR_ AS ALIAS_1 ON ALIAS_0.TIPOVALOR_ID = ALIAS_1.ID LEFT OUTER JOIN " & _
                          "V_COMPVALORES_ AS ALIAS_3 ON ALIAS_1.REFERENCIATIPOVALOR_ID = ALIAS_3.ID LEFT OUTER JOIN " & _
                          "V_UNIDADFINANCIERA_ AS ALIAS_2 ON ALIAS_0.VALOR2_UNIDADVALORIZACION_ID = ALIAS_2.ID LEFT OUTER JOIN " & _
                          "V_ESTADO_ AS ALIAS_4 ON ALIAS_0.ESTADO_ID = ALIAS_4.ID " & _
                          "WHERE     (ALIAS_0.BO_PLACE_ID IS NOT NULL) AND (ALIAS_3.DESCRIPCION NOT IN ('CHEQUE PROPIO', 'CHEQUE DIFERIDO PROPIO', 'DOCUMENTO A PAGAR')) AND " & _
                          "(ALIAS_1.TIPOVALOR LIKE '%cheque%') and V_CLIENTE_.ID = '" & datagrid1.Columns(0).Text & "'  " & _
                          "ORDER BY ALIAS_0.FECHAVEN DESC "

datvalores.Refresh
DataGrid2.Columns(0).Width = 3000
DataGrid2.Columns(1).Width = 1200
DataGrid2.Columns(2).Width = 1200
DataGrid2.Columns(3).Alignment = dbgRight
DataGrid2.Columns(3).NumberFormat = "#,###,##0.00"
DataGrid2.Columns(4).Alignment = 800
DataGrid2.Columns(6).Visible = False
DataGrid2.Columns(7).Visible = False




End Sub

Private Sub salir_Click()
    
    Unload Me

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

On Error Resume Next

   datparametros.ConnectionString = login.conexiontotal
   datparametros.RecordSource = "select * from ud_ezi_parametros_pos where sucursal = '" & login.nomsucursal & "' "
   datparametros.Refresh
   
   xcodclientefiltra = datparametros.Recordset.Fields("codclientefiltra")
   If xcodclientefiltra = "01" Then xcontrocliente = "88447B8E-14FE-4D60-9622-B22F6C735701"  ' tucuman
   If xcodclientefiltra = "04" Then xcontrocliente = "4234CA46-B2BE-4690-AC6A-F0DE206F94A9"  ' salta
   If xcodclientefiltra = "03" Then xcontrocliente = "AEC7FBAC-63F7-4404-9512-033D0961D9BC"  ' jujuy


    If KeyAscii = 13 Then
        KeyAscii = 0
        If Text1.Text <> "" Then
            Text1.Text = Replace(Text1.Text, " ", "%%")
            xbusqueda = "%" + Text1.Text + "%"
              
            xquery1 = "SELECT ALIAS_0.ID, ALIAS_0.CODIGO, ALIAS_3.NOMBRE AS RAZONSOCIAL, ALIAS_3.CUIT, ALIAS_6.CALLE, ISNULL(ALIAS_6.CALLE, '') " & _
              "+ '-' + ISNULL(V_CIUDAD_.NOMBRE, '') + '-' + ISNULL(V_PROVINCIA_.NOMBRE, '') AS DOMICILIO, ALIAS_0.DENOMINACION, ALIAS_5.NOMBRE AS ZONA, " & _
              "ALIAS_7.NUMERO AS TELEFONO, ALIAS_8.DIRECCIONELECTRONICA AS MAIL, V_TIPOPAGO_.NOMBRE AS TP, V_TIPOPAGO_.OBSERVACION AS TipoPago, " & _
              "V_TIPOPAGO_.ACTIVESTATUS AS ACTIVESTATUSTP, V_PROVINCIA_.NOMBRE AS Provincia, V_CIUDAD_.NOMBRE AS Ciudad, " & _
              "ALIAS_0.CODIGO + ' ' + ALIAS_3.CUIT + ' ' + ALIAS_3.NOMBRE AS CONCATENADO, V_EZI_POCICION_IVA_CLIENTES.CODIGO AS IVA, V_TIPOPAGO_.ID AS IDPAGO, " & _
              "v_ezi_pos_impuestos.COEFICIENTE AS Alic_IIBB, v_ezi_pos_impuestos.NOMBRE AS Cond_IIBB, v_ezi_pos_impuestos.EXENCION AS Exencion_IIBB,  V_UD_CLIENTE_.Anotaciones " & _
              "FROM         V_EZI_POCICION_IVA_CLIENTES RIGHT OUTER JOIN " & _
              "V_CLIENTE AS ALIAS_0 WITH (nolock) LEFT OUTER JOIN " & _
              "V_UD_CLIENTE_ ON ALIAS_0.BOEXTENSION_ID = V_UD_CLIENTE_.ID LEFT OUTER JOIN " & _
              "v_ezi_pos_impuestos ON ALIAS_0.ID = v_ezi_pos_impuestos.idcliente ON V_EZI_POCICION_IVA_CLIENTES.idcliente = ALIAS_0.ID LEFT OUTER JOIN " & _
              "V_TIPOPAGO_ ON ALIAS_0.TIPOPAGO_ID = V_TIPOPAGO_.ID LEFT OUTER JOIN " & _
              "V_PERSONA AS ALIAS_3 WITH (nolock) ON ALIAS_0.ENTEASOCIADO_ID = ALIAS_3.ID LEFT OUTER JOIN " & _
              "V_CIUDAD_ RIGHT OUTER JOIN " & _
              "V_DOMICILIO_ AS ALIAS_6 ON V_CIUDAD_.ID = ALIAS_6.CIUDAD_ID LEFT OUTER JOIN " & _
              "V_PROVINCIA_ ON ALIAS_6.PROVINCIA_ID = V_PROVINCIA_.ID ON ALIAS_0.DOMICILIOFACTURACION_ID = ALIAS_6.ID LEFT OUTER JOIN " & _
              "V_TELEFONO_ AS ALIAS_7 ON ALIAS_3.TELEFONOPRINCIPAL_ID = ALIAS_7.ID LEFT OUTER JOIN " & _
              "V_DIRECCIONELECTRONICA_ AS ALIAS_8 ON ALIAS_3.DIRECELECTRONICAPRINCIPAL_ID = ALIAS_8.ID LEFT OUTER JOIN " & _
              "V_ZONA_ AS ALIAS_5 ON ALIAS_0.ZONA_ID = ALIAS_5.ID " & _
              "WHERE     (ALIAS_0.ACTIVESTATUS = 0) AND ALIAS_0.CODIGO + ' ' + ALIAS_3.CUIT + ' ' + ALIAS_3.NOMBRE + ' ' + ALIAS_0.DENOMINACION   like '" & xbusqueda & "' " & _
              "order by ALIAS_3.NOMBRE "
              
              
                      
            datcliente.RecordSource = xquery1
            datcliente.Refresh
            datagrid1.Columns(0).Visible = False
            datagrid1.Columns(7).Visible = False
            'DataGrid1.Columns(9).Visible = False
            datagrid1.Columns(4).Visible = False
            datagrid1.Columns(10).Visible = False
            datagrid1.Columns(12).Visible = False
            datagrid1.Columns(15).Visible = False
            datagrid1.Columns(17).Visible = False
            datagrid1.Columns(18).Visible = False
            datagrid1.Columns(1).Width = 1000
            datagrid1.Columns(2).Width = 3500
            datagrid1.Columns(6).Width = 3500


        End If
        datagrid1.SetFocus
        
        
    End If

End Sub
