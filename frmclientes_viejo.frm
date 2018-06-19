VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form frmclientes_viejo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clientes"
   ClientHeight    =   8160
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   7845
   Icon            =   "frmclientes_viejo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8160
   ScaleWidth      =   7845
   Begin VB.CommandButton Command5 
      Caption         =   "Cod.Contable"
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
      Left            =   240
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Pers.Contacto"
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
      Left            =   240
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "E-mail"
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
      Left            =   240
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Teléfono"
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
      Index           =   8
      Left            =   240
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cod. Postal"
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
      Left            =   240
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Localidad"
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
      Left            =   240
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Domicilio"
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
      Left            =   240
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
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
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Tipo de Iva"
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
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Tipo de Cliente"
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
      Left            =   240
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Nombre o &Raz.Soc."
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
      Left            =   240
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   360
      Width           =   1815
   End
   Begin MSDataListLib.DataList DataList2 
      Bindings        =   "frmclientes_viejo.frx":0442
      Height          =   2205
      Left            =   3000
      TabIndex        =   26
      Top             =   3960
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3889
      _Version        =   393216
      MatchEntry      =   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      BackColor       =   12640511
      ListField       =   "codigo"
      BoundColumn     =   "Cod Contable"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "frmclientes_viejo.frx":045B
      Height          =   315
      Left            =   2160
      TabIndex        =   28
      Top             =   4440
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      BackColor       =   12640511
      ListField       =   "razonsocial"
      Text            =   ""
   End
   Begin VB.CommandButton busca 
      Caption         =   "busca"
      Height          =   255
      Left            =   5640
      TabIndex        =   25
      Top             =   960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   255
      Left            =   840
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtFields 
      Alignment       =   2  'Center
      DataField       =   "codcontable"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   11
      Left            =   2280
      TabIndex        =   23
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox txtFields 
      DataField       =   "tipocliente"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   720
      Width           =   495
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frmclientes_viejo.frx":0476
      DataField       =   "tipoiva"
      DataSource      =   "datPrimaryRS"
      Height          =   315
      Left            =   2880
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   741
      _Version        =   393216
      Style           =   2
      BackColor       =   14737632
      ListField       =   "descripcion"
      BoundColumn     =   "categ"
      Text            =   "DataCombo1"
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
   Begin VB.CommandButton ordenarazosocial 
      Caption         =   "Razon Social"
      Height          =   255
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4560
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmclientes_viejo.frx":0490
      Height          =   3255
      Left            =   120
      TabIndex        =   16
      Top             =   4800
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   5741
      _Version        =   393216
      BackColor       =   14737632
      HeadLines       =   1
      RowHeight       =   15
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "codcliente"
         Caption         =   "Cod.Cliente"
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
      BeginProperty Column02 
         DataField       =   "tipocliente"
         Caption         =   "tipocliente"
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
      BeginProperty Column03 
         DataField       =   "razonsocial"
         Caption         =   "Nombre o Razon Social"
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
      BeginProperty Column04 
         DataField       =   "tipoiva"
         Caption         =   "tipoiva"
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
      BeginProperty Column05 
         DataField       =   "cuit"
         Caption         =   "cuit"
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
      BeginProperty Column06 
         DataField       =   "domicilio"
         Caption         =   "domicilio"
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
      BeginProperty Column07 
         DataField       =   "localidad"
         Caption         =   "localidad"
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
      BeginProperty Column08 
         DataField       =   "codpostal"
         Caption         =   "codpostal"
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
      BeginProperty Column09 
         DataField       =   "telefono"
         Caption         =   "telefono"
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
      BeginProperty Column10 
         DataField       =   "email"
         Caption         =   "email"
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
      BeginProperty Column11 
         DataField       =   "contacto"
         Caption         =   "contacto"
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
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column07 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column08 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column09 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column10 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column11 
            Locked          =   -1  'True
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtFields 
      DataField       =   "contacto"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   10
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   9
      Top             =   3600
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "email"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   9
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   8
      Top             =   3240
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "telefono"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   8
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   7
      Top             =   2880
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      Alignment       =   2  'Center
      DataField       =   "codpostal"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   7
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   6
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txtFields 
      DataField       =   "localidad"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   6
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   5
      Top             =   2160
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "domicilio"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   5
      Left            =   2280
      MaxLength       =   40
      TabIndex        =   4
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "tipoiva"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   4
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox txtFields 
      DataField       =   "razonsocial"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   2
      Top             =   345
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "empresa"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtFields 
      DataField       =   "codcliente"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   7830
      Visible         =   0   'False
      Width           =   7845
      _ExtentX        =   13838
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
      Caption         =   " "
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
   Begin VB.CommandButton borrar 
      Caption         =   "&Borrar"
      Height          =   615
      Left            =   6480
      Picture         =   "frmclientes_viejo.frx":04AB
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton grabar 
      Caption         =   "&Grabar"
      Height          =   615
      Left            =   6480
      Picture         =   "frmclientes_viejo.frx":05AD
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton nuevo 
      Caption         =   "&Nuevo"
      Height          =   615
      Left            =   6480
      Picture         =   "frmclientes_viejo.frx":0ADF
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton Cancelar 
      Cancel          =   -1  'True
      Caption         =   "Cance&lar"
      Height          =   615
      Left            =   6480
      Picture         =   "frmclientes_viejo.frx":1011
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton salir 
      Caption         =   "&Cerrar"
      Height          =   615
      Left            =   6480
      Picture         =   "frmclientes_viejo.frx":1543
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3960
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin MSAdodcLib.Adodc datacontrib 
      Height          =   330
      Left            =   4560
      Top             =   5160
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
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   5775
      Begin MSDataListLib.DataCombo tipcliente 
         Bindings        =   "frmclientes_viejo.frx":1985
         Height          =   315
         Left            =   2760
         TabIndex        =   27
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   741
         _Version        =   393216
         MatchEntry      =   -1  'True
         BackColor       =   14737632
         ListField       =   "tipoclientes"
         BoundColumn     =   "codigo"
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
      Begin MSMask.MaskEdBox maskcuit 
         DataField       =   "cuit"
         DataSource      =   "datPrimaryRS"
         Height          =   255
         Left            =   2160
         TabIndex        =   40
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   13
         PromptChar      =   "_"
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
      Height          =   4575
      Left            =   6120
      TabIndex        =   18
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton Command1 
         Caption         =   "&Listar"
         Height          =   615
         Left            =   360
         Picture         =   "frmclientes_viejo.frx":19A3
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   3120
         UseMaskColor    =   -1  'True
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      Picture         =   "frmclientes_viejo.frx":1AA5
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   20
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton ordenarcodigo 
      Caption         =   "Cod.Clientes"
      Height          =   255
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin Crystal.CrystalReport CrystalReporte 
      Left            =   5160
      Top             =   -120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Libro IVA Compras"
      PrintFileLinesPerPage=   60
   End
   Begin MSRDC.MSRDC reporte 
      Height          =   375
      Left            =   3840
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
   Begin MSAdodcLib.Adodc datbusca 
      Height          =   330
      Left            =   6600
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
   Begin vbskpro.Skinner Skinner1 
      Left            =   5640
      Top             =   -120
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
      LcK2            =   $"frmclientes_viejo.frx":229F
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
   Begin MSAdodcLib.Adodc datcuentas 
      Height          =   330
      Left            =   6600
      Top             =   4320
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
Attribute VB_Name = "frmclientes_viejo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim empresareal
Private Sub borrar_Click()
On Error GoTo errorborrado

KeyAscii = 13
  respuesta = MsgBox("ESTA POR BORRAR UN CLIENTE, ESTA SEGURO?", vbYesNo, "Atención")
If respuesta = vbYes Then

                Inicio.datauditoria.ConnectionString = login.conexiontotal
    
                Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
                Inicio.datauditoria.Refresh
    
                Inicio.datauditoria.Recordset.AddNew
                Inicio.datauditoria.Recordset.Fields("fecha") = Date
                Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
                Inicio.datauditoria.Recordset.Fields("ventana") = "CLIENTES"
                Inicio.datauditoria.Recordset.Fields("accion") = "Eliminacion Cliente:" + txtFields(2).Text
                Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
                Inicio.datauditoria.Recordset.Fields("empresa") = empresareal
                Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent

    datPrimaryRS.Recordset.Delete
Else
    Exit Sub
End If

Exit Sub
errorborrado:

    MsgBox ("No se pudo borrar el registro")


End Sub

Private Sub busca_Click()
On Error GoTo fuera

    datbusca.RecordSource = "select clientes.* from clientes where empresa = " & empresareal & " and cuit = '" & maskcuit & "'"
    datbusca.Refresh
    
    If datbusca.Recordset.EOF = True Then
        datbusca.RecordSource = "select clientes.* from clientes where empresa = " & empresareal & " "
        datbusca.Refresh
        Exit Sub
    Else
        mensa = MsgBox("Este Cliente ya fue ingresado", vbCritical, "!! Atención !!")
    Rem    Call Cancelar_Click
        datbusca.RecordSource = "select clientes.* from clientes where empresa = " & empresareal & " "
        datbusca.Refresh
    End If
    
fuera:
End Sub

Private Sub cancelar_Click()

    datPrimaryRS.Refresh

End Sub



Private Sub Command1_Click()
On Error GoTo fuera

                Inicio.datauditoria.ConnectionString = login.conexiontotal
    
                Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
                Inicio.datauditoria.Refresh
    
                Inicio.datauditoria.Recordset.AddNew
                Inicio.datauditoria.Recordset.Fields("fecha") = Date
                Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
                Inicio.datauditoria.Recordset.Fields("ventana") = "CLIENTES"
                Inicio.datauditoria.Recordset.Fields("accion") = "Listado Clientes"
                Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
                Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
                Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent

    Call Command2_Click

fuera:
End Sub

Private Sub Command2_Click()
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

With crystalreporte
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

Private Sub DataCombo1_Click(Area As Integer)
On Error GoTo fuera

    txtFields(4).Text = DataCombo1.BoundText

fuera:
End Sub

Private Sub DataCombo1_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Then
        KeyAscii = 0
        maskcuit.SetFocus
    End If
    
fuera:
End Sub



Private Sub DataCombo2_Click(Area As Integer)
On Error GoTo fuera
If DataCombo2.Text <> "" Then
    DataGrid1.Bookmark = DataCombo2.SelectedItem
End If
fuera:
End Sub

Private Sub DataCombo2_KeyPress(KeyAscii As Integer)
On Error GoTo fuera
    If KeyAscii = 13 Then
        KeyAscii = 0
        If DataCombo2.Text <> "" Then
            DataGrid1.Bookmark = DataCombo2.SelectedItem
        End If
    End If
fuera:
End Sub

Private Sub DataCombo2_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo fuera
If DataCombo2.Text <> "" Then
    DataGrid1.Bookmark = DataCombo2.SelectedItem
End If
fuera:
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)

tipcliente.BoundText = txtFields(3).Text

End Sub

Private Sub DataList2_GotFocus()
On Error GoTo fuera

    If Inicio.opcion1 = True Then
        datcuentas.RecordSource = "select listacuentas.* from listacuentas WHERE empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' ORDER BY IDCUENTA"
        datcuentas.Refresh
        DataList2.ListField = "codigo"
    End If
    If Inicio.opcion2 = True Then
        datcuentas.RecordSource = "select listacuentas.* from listacuentas WHERE empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' ORDER BY nombre"
        datcuentas.Refresh
        DataList2.ListField = "nombre"
    End If

fuera:
End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Then
        KeyAscii = 0
        If DataList2.Text <> "" Then txtFields(11).Text = DataList2.BoundText
        DataList2.Visible = False
        grabar.SetFocus
    End If

fuera:
End Sub

Private Sub DataList2_LostFocus()

    DataList2.Visible = False

End Sub

Private Sub Form_GotFocus()
        maskcuit.Mask = ""
        maskcuit.MaxLength = 13
End Sub

Private Sub Form_Load()
Aplicar_skin Me

frmclientes.Top = 0

    Inicio.Toolbar1.Visible = True
    
datbusca.ConnectionString = login.conexiontotal
datcuentas.ConnectionString = login.conexiontotal
datPrimaryRS.ConnectionString = login.conexiontotal
datacontrib.ConnectionString = login.conexiontotal
dattipoclientes.ConnectionString = login.conexiontotal

If login.empresaact1 > 0 Then
    empresareal = login.empresaact1
Else
    empresareal = login.empresaact
End If

    dattipoclientes.RecordSource = "select tipoclientes.* from tipoclientes where empresa = " & empresareal & ""
    dattipoclientes.Refresh
    datacontrib.RecordSource = "select condtrib.* from condtrib"
    datacontrib.Refresh
     
    datPrimaryRS.RecordSource = "select clientes.* from clientes where clientes.empresa = " & empresareal & " ORDER BY razonsocial"
    datPrimaryRS.Refresh
    datcuentas.RecordSource = "select cuentas.* from cuentas where empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' and imp = 'S' ORDER BY IDCUENTA"
    datcuentas.Refresh
    
    DataCombo2.Text = txtFields(2).Text
    
    maskcuit.Mask = "##-########-#"
    maskcuit.MaxLength = 13
    If datPrimaryRS.Recordset.EOF = True Then
            datPrimaryRS.Recordset.AddNew
            txtFields(1) = empresareal
            txtFields(3) = 1
            maskcuit.SelLength = 13
            maskcuit.SelText = ""
            
    End If
                       
 Rem   tipcliente.AddItem ("COMERCIAL")
 Rem   tipcliente.AddItem ("OFICIAL")
 Rem   tipcliente.AddItem ("NACIONAL")
 Rem   tipcliente.Text = tipcliente.List(Val(txtFields(3).Text) - 1)
        
 Rem      tipcliente.SelectedItem = txtFields(3).Text
      tipcliente.BoundText = txtFields(3).Text
                       
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


Private Sub grabar_Click()
On Error GoTo fuera

     If txtFields(5).Text = "" Then txtFields(5).Text = "-"
     datPrimaryRS.Recordset.UpdateBatch adAffectCurrent
     datPrimaryRS.Refresh
     
                Inicio.datauditoria.ConnectionString = login.conexiontotal
    
                Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
                Inicio.datauditoria.Refresh
    
                Inicio.datauditoria.Recordset.AddNew
                Inicio.datauditoria.Recordset.Fields("fecha") = Date
                Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
                Inicio.datauditoria.Recordset.Fields("ventana") = "CLIENTES"
                Inicio.datauditoria.Recordset.Fields("accion") = "Modif.Cliente:" + txtFields(2).Text
                Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
                Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
                Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent
     
     nuevo.SetFocus

fuera:
End Sub


Private Sub maskcuit_Change()
On Error GoTo fuera

    If KeyAscii = 13 Then
        KeyAscii = 0
        KeyAscii = 9
    End If

fuera:
End Sub

Private Sub maskcuit_LostFocus()


    Call busca_Click
    mensa = verifica_cuit(maskcuit.Text)

End Sub

Private Sub nuevo_Click()
On Error GoTo fuera

    datPrimaryRS.Recordset.AddNew
    txtFields(1) = empresareal
    maskcuit.SelLength = 13
    maskcuit.SelText = ""
    txtFields(3).Text = 1
    
                Inicio.datauditoria.ConnectionString = login.conexiontotal
    
                Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
                Inicio.datauditoria.Refresh
    
                Inicio.datauditoria.Recordset.AddNew
                Inicio.datauditoria.Recordset.Fields("fecha") = Date
                Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
                Inicio.datauditoria.Recordset.Fields("ventana") = "CLIENTES"
                Inicio.datauditoria.Recordset.Fields("accion") = "Alta cliente"
                Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
                Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
                Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent
    
    
    txtFields(2).SetFocus

fuera:
End Sub

Private Sub ordenarazosocial_Click()

    datPrimaryRS.RecordSource = "select clientes.* from clientes  WHERE clientes.empresa = " & empresareal & " ORDER BY razonsocial"
    datPrimaryRS.Refresh
    
End Sub

Private Sub ordenarcodigo_Click()

    datPrimaryRS.RecordSource = "select clientes.* from clientes WHERE clientes.empresa = " & empresareal & " ORDER BY codcliente"
    datPrimaryRS.Refresh
    
End Sub

Private Sub salir_Click()

    Unload Me

End Sub


Private Sub tipcliente_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Then
        KeyAscii = 0
        txtFields(3).Text = tipcliente.BoundText
        DataCombo1.SetFocus
    End If

fuera:
End Sub

Private Sub txtFields_Change(Index As Integer)
On Error GoTo fuera

 Rem    tipcliente.Text = tipcliente.List(Val(txtFields(3).Text) - 1)

fuera:
End Sub

Private Sub txtFields_GotFocus(Index As Integer)
On Error GoTo fuera

If Index = 11 Then
    DataList2.Visible = True
    DataList2.SetFocus
End If

fuera:
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Then
        KeyAscii = 0
        If Index = 2 Then
            tipcliente.SetFocus
            Exit Sub
        End If
        If Index = 4 Then
            maskcuit.SetFocus
            Exit Sub
        End If
        If Index = 11 Then
            grabar.SetFocus
            Exit Sub
        End If
        txtFields(Index + 1).SetFocus
    End If
    
fuera:
End Sub
