VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form frmccostos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Centros de Costo"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9045
   Icon            =   "frmccostos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Palette         =   "frmccostos.frx":030A
   ScaleHeight     =   8640
   ScaleWidth      =   9045
   Begin VB.CommandButton actualiza 
      Caption         =   "actualiza"
      Height          =   255
      Left            =   7560
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame4 
      Caption         =   "Nievel 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   960
      TabIndex        =   6
      Top             =   6600
      Width           =   7695
      Begin MSDataGridLib.DataGrid DataGrid4 
         Bindings        =   "frmccostos.frx":0494
         Height          =   1215
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2143
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   2
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
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "idcc"
            Caption         =   "idcc"
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
            DataField       =   "idsubcc1"
            Caption         =   "idsubcc1"
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
            DataField       =   "idsubcc2"
            Caption         =   "idsubcc2"
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
            DataField       =   "subcc3"
            Caption         =   "Codigo"
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
            DataField       =   "descripcion"
            Caption         =   "Descripcion"
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
            DataField       =   "porcentaje"
            Caption         =   "%"
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
               Locked          =   -1  'True
               Object.Visible         =   0   'False
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
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Nuevo"
         Height          =   615
         Left            =   6480
         Picture         =   "frmccostos.frx":04A9
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Eliminar"
         Height          =   615
         Left            =   6480
         Picture         =   "frmccostos.frx":09DB
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Nievel 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   720
      TabIndex        =   4
      Top             =   4800
      Width           =   7935
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "frmccostos.frx":0B65
         Height          =   1215
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2143
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   2
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "idcc"
            Caption         =   "idcc"
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
            DataField       =   "idsubcc1"
            Caption         =   "idsubcc1"
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
            DataField       =   "subcc2"
            Caption         =   "Codigo"
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
            DataField       =   "descripcion"
            Caption         =   "Descripcion"
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
            DataField       =   "porcentaje"
            Caption         =   "%"
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
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
            EndProperty
            BeginProperty Column05 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Nuevo"
         Height          =   615
         Left            =   6720
         Picture         =   "frmccostos.frx":0B7A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Eliminar"
         Height          =   615
         Left            =   6720
         Picture         =   "frmccostos.frx":10AC
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Nievel 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   480
      TabIndex        =   2
      Top             =   3000
      Width           =   8175
      Begin VB.CommandButton Command2 
         Caption         =   "Nuevo"
         Height          =   615
         Left            =   6960
         Picture         =   "frmccostos.frx":1236
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Eliminar"
         Height          =   615
         Left            =   6960
         Picture         =   "frmccostos.frx":1768
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmccostos.frx":18F2
         Height          =   1215
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2143
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   2
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "idcc"
            Caption         =   "idcc"
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
            DataField       =   "subcc1"
            Caption         =   "Codigo"
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
            DataField       =   "descripcion"
            Caption         =   "Descripcion"
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
            DataField       =   "porcentaje"
            Caption         =   "%"
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
               Alignment       =   1
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   -1  'True
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nievel 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8415
      Begin VB.CommandButton eliminar 
         Caption         =   "Eliminar"
         Height          =   615
         Left            =   7080
         Picture         =   "frmccostos.frx":1907
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton nuevo 
         Caption         =   "Nuevo"
         Height          =   615
         Left            =   7080
         Picture         =   "frmccostos.frx":1A91
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmccostos.frx":1FC3
         Height          =   2175
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   3836
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         BackColor       =   12648447
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   2
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "cc"
            Caption         =   "Codigo"
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
            DataField       =   "descripcion"
            Caption         =   "Descripcion"
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
         BeginProperty Column03 
            DataField       =   "digitocontable"
            Caption         =   "1º Dig.Egresos"
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
            DataField       =   "digitocontable1"
            Caption         =   "1º Dig.Ingresos"
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
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc datcc 
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
   Begin vbskpro.Skinner Skinner1 
      Left            =   6240
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
      LcK2            =   $"frmccostos.frx":1FD7
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
   Begin MSAdodcLib.Adodc datcc1 
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select ccostos1.* from ccostos1"
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
   Begin MSAdodcLib.Adodc datcc2 
      Height          =   330
      Left            =   2640
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select ccostos2.* from ccostos2"
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
   Begin MSAdodcLib.Adodc datcc3 
      Height          =   330
      Left            =   3960
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select ccostos3.* from ccostos3"
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
   Begin VB.Frame Frame5 
      Height          =   1095
      Left            =   360
      TabIndex        =   17
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Frame Frame6 
      Height          =   1095
      Left            =   600
      TabIndex        =   18
      Top             =   4560
      Width           =   1215
      Begin VB.Frame Frame7 
         Height          =   1095
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Frame Frame8 
      Height          =   1095
      Left            =   840
      TabIndex        =   20
      Top             =   6360
      Width           =   1215
   End
End
Attribute VB_Name = "frmccostos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nivel As Integer
Private Sub actualiza_Click()
On Error GoTo fuera
If nivel = 1 Then
    datcc1.RecordSource = "select ccostos1.* from ccostos1 where idcc = " & DataGrid1.Columns(0) & ""
    datcc1.Refresh
End If

If nivel <= 2 Then
    If datcc1.Recordset.EOF = False Then
        datcc2.RecordSource = "select ccostos2.* from ccostos2 where idcc = " & DataGrid1.Columns(0) & " and idsubcc1 = " & DataGrid2.Columns(1) & ""
        datcc2.Refresh
        Frame3.Caption = DataGrid1.Columns(1) + " - " + DataGrid2.Columns(2)
        If datcc2.Recordset.EOF = False Then
            datcc3.RecordSource = "select ccostos3.* from ccostos3 where idcc = " & DataGrid1.Columns(0) & " and idsubcc1 = " & DataGrid2.Columns(1) & " and idsubcc2 = " & DataGrid3.Columns(2) & ""
            datcc3.Refresh
            Frame4.Caption = DataGrid1.Columns(1) + " - " + DataGrid2.Columns(2) + " - " + DataGrid3.Columns(3)
        End If
    Else
        datcc2.RecordSource = "select ccostos2.* from ccostos2 where idcc = 0"
        datcc2.Refresh
        Frame3.Caption = ""
        Frame4.Caption = ""
    End If
End If
    
If datcc2.Recordset.EOF = False Then
    If nivel = 3 Then
        datcc3.RecordSource = "select ccostos3.* from ccostos3 where idcc = " & DataGrid1.Columns(0) & " and idsubcc1 = " & DataGrid2.Columns(1) & " and idsubcc2 = " & DataGrid3.Columns(2) & ""
        datcc3.Refresh
        Frame4.Caption = DataGrid1.Columns(1) + " - " + DataGrid2.Columns(2) + " - " + DataGrid3.Columns(3)
    End If
Else
        datcc3.RecordSource = "select ccostos3.* from ccostos3 where idcc = 0"
        datcc3.Refresh
        Frame4.Caption = ""
End If
    
    
    Frame2.Caption = DataGrid1.Columns(1)
fuera:
End Sub

Private Sub Command1_Click()
On Error GoTo fuera

    mensa = MsgBox("Esta por borrar un Sub centro de costo, esta seguro", vbYesNo, "!Atencion!")
    If mensa = vbYes Then
        datcc1.Recordset.Delete adAffectCurrent
        datcc1.Refresh
        Exit Sub
    End If

fuera:
End Sub

Private Sub Command2_Click()

On Error GoTo fuera

    datcc1.Recordset.AddNew
    datcc1.Recordset.Fields(4).Value = login.empresaact
    datcc1.Recordset.Fields(0).Value = DataGrid1.Columns(0)
    DataGrid2.Col = 1
    DataGrid2.SetFocus
    
fuera:


End Sub

Private Sub Command3_Click()
On Error GoTo fuera

    mensa = MsgBox("Esta por borrar un Sub centro de costo, esta seguro", vbYesNo, "!Atencion!")
    If mensa = vbYes Then
        datcc2.Recordset.Delete adAffectCurrent
        datcc2.Refresh
        Exit Sub
    End If

fuera:
End Sub

Private Sub Command4_Click()

On Error GoTo fuera

    datcc2.Recordset.AddNew
    datcc2.Recordset.Fields(5).Value = login.empresaact
    datcc2.Recordset.Fields(0).Value = DataGrid1.Columns(0)
    datcc2.Recordset.Fields(1).Value = DataGrid2.Columns(1)
    DataGrid3.Col = 2
    DataGrid3.SetFocus
    
fuera:


End Sub

Private Sub Command5_Click()
On Error GoTo fuera

    mensa = MsgBox("Esta por borrar un Sub centro de costo, esta seguro", vbYesNo, "!Atencion!")
    If mensa = vbYes Then
        datcc3.Recordset.Delete adAffectCurrent
        datcc3.Refresh
        Exit Sub
    End If

fuera:
End Sub

Private Sub Command6_Click()
On Error GoTo fuera

    datcc3.Recordset.AddNew
    datcc3.Recordset.Fields(6).Value = login.empresaact
    datcc3.Recordset.Fields(0).Value = DataGrid1.Columns(0)
    datcc3.Recordset.Fields(1).Value = DataGrid2.Columns(1)
    datcc3.Recordset.Fields(2).Value = DataGrid3.Columns(2)
    DataGrid4.Col = 3
    DataGrid4.SetFocus
    
fuera:
End Sub

Private Sub DataGrid1_Click()

    nivel = 1
    Call actualiza_Click
    

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Then
        KeyAscii = 0
        KeyAscii = 9
    End If
    
fuera:
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)

    nivel = 1
    Call actualiza_Click

End Sub

Private Sub DataGrid2_AfterColEdit(ByVal ColIndex As Integer)
On Error GoTo fuera
        datcc1.Recordset.UpdateBatch adAffectCurrent
fuera:
End Sub

Private Sub DataGrid2_Click()

    nivel = 2
    Call actualiza_Click


End Sub

Private Sub DataGrid2_KeyPress(KeyAscii As Integer)

On Error GoTo fuera

    If KeyAscii = 13 Then
        KeyAscii = 0
        KeyAscii = 9
    End If
    
fuera:

End Sub

Private Sub DataGrid2_KeyUp(KeyCode As Integer, Shift As Integer)

    nivel = 2
    Call actualiza_Click

End Sub

Private Sub DataGrid3_AfterColEdit(ByVal ColIndex As Integer)
On Error GoTo fuera
datcc2.Recordset.UpdateBatch adAffectCurrent
fuera:
End Sub

Private Sub DataGrid3_Click()

    nivel = 3
    Call actualiza_Click

End Sub

Private Sub DataGrid3_KeyPress(KeyAscii As Integer)

On Error GoTo fuera

    If KeyAscii = 13 Then
        KeyAscii = 0
        KeyAscii = 9
    End If
    
fuera:

End Sub

Private Sub DataGrid3_KeyUp(KeyCode As Integer, Shift As Integer)
    nivel = 3
    Call actualiza_Click
End Sub

Private Sub DataGrid4_AfterColEdit(ByVal ColIndex As Integer)
On Error GoTo fuera
datcc3.Recordset.UpdateBatch adAffectCurrent
fuera:
End Sub

Private Sub DataGrid4_KeyPress(KeyAscii As Integer)
On Error GoTo fuera

    If KeyAscii = 13 Then
        KeyAscii = 0
        KeyAscii = 9
    End If
    
fuera:
End Sub

Private Sub eliminar_Click()
On Error GoTo fuera

    mensa = MsgBox("Esta por borrar un centro de costo, esta seguro", vbYesNo, "!Atencion!")
    If mensa = vbYes Then
        datcc.Recordset.Delete adAffectCurrent
        datcc.Refresh
        Exit Sub
    End If

fuera:
End Sub

Private Sub Form_Load()
On Error GoTo fuera

datcc.ConnectionString = login.conexiontotal
Rem datcc1.ConnectionString = login.conexiontotal
Rem datcc2.ConnectionString = login.conexiontotal
Rem datcc3.ConnectionString = login.conexiontotal

    datcc.RecordSource = "select ccostos.* from ccostos where empresa = " & login.empresaact & " order by cc "
    datcc.Refresh
    nivel = 1
    Call actualiza_Click

fuera:
End Sub

Private Sub nuevo_Click()
On Error GoTo fuera

    datcc.Recordset.AddNew
    datcc.Recordset.Fields(2).Value = login.empresaact
    DataGrid1.SetFocus
    
fuera:
End Sub
