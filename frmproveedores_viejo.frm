VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form frmproveedores_viejo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proveedores"
   ClientHeight    =   8565
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   8505
   Icon            =   "frmproveedores_viejo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   8505
   Begin MSDataListLib.DataList DataList2 
      Bindings        =   "frmproveedores_viejo.frx":0442
      Height          =   2205
      Left            =   3000
      TabIndex        =   38
      Top             =   4200
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmproveedores_viejo.frx":045B
      Height          =   3255
      Left            =   120
      TabIndex        =   50
      Top             =   5160
      Width           =   8175
      _ExtentX        =   14420
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
         DataField       =   "codproveedor"
         Caption         =   "Cod.Proveedor"
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
   Begin VB.CommandButton ordenarazosocial 
      Caption         =   "Razon Social"
      Height          =   255
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   4920
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton ordenarcodigo 
      Caption         =   "Cod.Proved."
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   4920
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin MSDataListLib.DataList DataList1 
      Bindings        =   "frmproveedores_viejo.frx":0476
      Height          =   2205
      Left            =   3360
      TabIndex        =   47
      Top             =   1680
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
   Begin VB.TextBox txtFields 
      DataField       =   "codproveedor"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   8070
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos &1"
      TabPicture(0)   =   "frmproveedores_viejo.frx":048F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblLabels(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblLabels(10)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblLabels(9)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblLabels(8)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblLabels(7)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblLabels(6)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblLabels(5)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblLabels(4)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblLabels(3)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblLabels(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "maskcuit"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "DataCombo1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Check1(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtFields(11)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtFields(10)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtFields(9)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtFields(8)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtFields(7)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtFields(6)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtFields(5)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtFields(3)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtFields(2)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "Datos &2"
      TabPicture(1)   =   "frmproveedores_viejo.frx":04AB
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text4"
      Tab(1).Control(1)=   "Text3"
      Tab(1).Control(2)=   "fechacai"
      Tab(1).Control(3)=   "lblLabels(13)"
      Tab(1).Control(4)=   "lblLabels(12)"
      Tab(1).Control(5)=   "lblLabels(11)"
      Tab(1).Control(6)=   "lblLabels(1)"
      Tab(1).ControlCount=   7
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         DataField       =   "codcontablegastos"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   -72600
         TabIndex        =   46
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         DataField       =   "plazovencfacturas"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   -72600
         TabIndex        =   43
         Top             =   1080
         Width           =   855
      End
      Begin Project1.jeffMaskedEdit fechacai 
         Height          =   255
         Left            =   -72600
         TabIndex        =   40
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         MouseIcon       =   "frmproveedores_viejo.frx":04C7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         SelText         =   ""
         Text            =   "__/__/____"
         HideSelection   =   -1  'True
         Alignment       =   2
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   -480
         Picture         =   "frmproveedores_viejo.frx":04E3
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   24
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "razonsocial"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   2
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   21
         Top             =   735
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "tipoiva"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   3
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1155
         Width           =   495
      End
      Begin VB.TextBox txtFields 
         DataField       =   "domicilio"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   5
         Left            =   2160
         MaxLength       =   40
         TabIndex        =   19
         Top             =   1935
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "localidad"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   6
         Left            =   2160
         MaxLength       =   30
         TabIndex        =   18
         Top             =   2295
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         Alignment       =   2  'Center
         DataField       =   "codpostal"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   7
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   17
         Top             =   2655
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         DataField       =   "telefono"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   8
         Left            =   2160
         MaxLength       =   30
         TabIndex        =   16
         Top             =   3015
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "email"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   9
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   15
         Top             =   3375
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "contacto"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   10
         Left            =   2160
         MaxLength       =   30
         TabIndex        =   14
         Top             =   3735
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         Alignment       =   2  'Center
         DataField       =   "codcontable"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   11
         Left            =   2160
         TabIndex        =   13
         Top             =   4095
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "Cont.:"
         Top             =   1575
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   0
         Left            =   4440
         TabIndex        =   11
         Top             =   1575
         Width           =   255
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmproveedores_viejo.frx":0CDD
         DataField       =   "tipoiva"
         DataSource      =   "datPrimaryRS"
         Height          =   315
         Left            =   2640
         TabIndex        =   22
         Top             =   1125
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   14737632
         ListField       =   "descripcion"
         BoundColumn     =   "categ"
         Text            =   "DataCombo1"
      End
      Begin MSMask.MaskEdBox maskcuit 
         Bindings        =   "frmproveedores_viejo.frx":0CF7
         DataField       =   "cuit"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   2160
         TabIndex        =   23
         Top             =   1575
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Frame Frame1 
         Height          =   3960
         Left            =   2040
         TabIndex        =   25
         Top             =   495
         Width           =   4215
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   255
            Index           =   1
            Left            =   3360
            TabIndex        =   26
            Top             =   1080
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   27
            TabStop         =   0   'False
            Text            =   "Cta.Cte.:"
            Top             =   1080
            Width           =   735
         End
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cod.Cont. Gastos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   -74640
         TabIndex        =   45
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Días"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   -71640
         TabIndex        =   44
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Plazo Venc. Fact.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   -74640
         TabIndex        =   42
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   -74640
         TabIndex        =   39
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   37
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   36
         Top             =   1155
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   35
         Top             =   1575
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   34
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   33
         Top             =   2295
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   32
         Top             =   2655
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   31
         Top             =   3015
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   30
         Top             =   3375
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   29
         Top             =   3750
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   4095
         Width           =   1815
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8760
      Top             =   5880
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
      DataSourceName  =   "contable"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select listacuentas.* from listacuentas ORDER BY IDCUENTA"
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
   Begin VB.CommandButton busca 
      Caption         =   "busca"
      Height          =   255
      Left            =   8280
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Listar"
      Height          =   615
      Left            =   7080
      Picture         =   "frmproveedores_viejo.frx":0D1B
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3600
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Cancelar 
      Cancel          =   -1  'True
      Caption         =   "Cance&lar"
      Height          =   615
      Left            =   7080
      Picture         =   "frmproveedores_viejo.frx":124D
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      DataField       =   "empresa"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   1920
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
      Top             =   8235
      Visible         =   0   'False
      Width           =   8505
      _ExtentX        =   15002
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
      Left            =   7080
      Picture         =   "frmproveedores_viejo.frx":177F
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton grabar 
      Caption         =   "&Grabar"
      Height          =   615
      Left            =   7080
      Picture         =   "frmproveedores_viejo.frx":1881
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton nuevo 
      Caption         =   "&Nuevo"
      Height          =   615
      Left            =   7080
      Picture         =   "frmproveedores_viejo.frx":1DB3
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton salir 
      Caption         =   "&Cerrar"
      Height          =   615
      Left            =   7080
      Picture         =   "frmproveedores_viejo.frx":22E5
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin MSAdodcLib.Adodc datacontrib 
      Height          =   330
      Left            =   4800
      Top             =   7080
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
   Begin MSAdodcLib.Adodc datcuentas 
      Height          =   330
      Left            =   0
      Top             =   6000
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
   Begin Crystal.CrystalReport CrystalReporte 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Libro IVA Compras"
      PrintFileLinesPerPage=   60
   End
   Begin MSRDC.MSRDC reporte 
      Height          =   375
      Left            =   3240
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
      Left            =   240
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
      LcK2            =   $"frmproveedores_viejo.frx":2727
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
      Height          =   4935
      Left            =   6720
      TabIndex        =   6
      Top             =   360
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "frmproveedores_viejo.frx":2736
      Height          =   315
      Left            =   3000
      TabIndex        =   51
      Top             =   4800
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      BackColor       =   12640511
      ListField       =   "razonsocial"
      Text            =   ""
   End
End
Attribute VB_Name = "frmproveedores_viejo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cuitprov As String
Public nomprove As String
Dim empresareal As Integer

Private Sub borrar_Click()
On Error GoTo errorborrado

KeyAscii = 13
  respuesta = MsgBox("ESTA POR BORRAR UN PROVEEDOR, ESTA SEGURO?", vbYesNo, "Atención")
If respuesta = vbYes Then

    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Proveedores"
    Inicio.datauditoria.Recordset.Fields("accion") = "Baja Prove.:" + txtFields(2).Text
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
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

    datbusca.RecordSource = "select proveedores.* from proveedores where empresa = " & empresareal & " and cuit = '" & maskcuit & "'  ORDER BY codproveedor"
    datbusca.Refresh
    
    If datbusca.Recordset.EOF = True Then
        datbusca.RecordSource = "select proveedores.* from proveedores where empresa = " & empresareal & "  ORDER BY codproveedor"
        datbusca.Refresh
        Exit Sub
    Else
        mensa = MsgBox("Este Proveedor ya fue ingresado", vbCritical, "!! Atención !!")
        Call cancelar_Click
        datbusca.RecordSource = "select proveedores.* from proveedores where empresa = " & empresareal & "  ORDER BY codproveedor"
        datbusca.Refresh
    End If
    
    

End Sub

Private Sub cancelar_Click()

  
     datPrimaryRS.Refresh
    

End Sub



Private Sub Check1_Click(Index As Integer)

If Index = 0 Then
    If Check1(0).Value = 1 Then
        Check1(1).Value = 0
    Else
        Check1(1).Value = 1
    End If
End If
If Index = 1 Then
    If Check1(1).Value = 1 Then
        Check1(0).Value = 0
    Else
        Check1(0).Value = 1
    End If
End If

End Sub

Private Sub Command1_Click()


    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Proveedores"
    Inicio.datauditoria.Recordset.Fields("accion") = "Listado Proveedores"
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent

    Call Command2_Click
End Sub

Private Sub Command2_Click()
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

Private Sub DataCombo1_Click(Area As Integer)

    txtFields(3).Text = DataCombo1.BoundText

End Sub

Private Sub DataCombo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        maskcuit.SetFocus
    End If
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

Private Sub DataGrid1_Click()

    If IsNull(datPrimaryRS.Recordset.Fields("fechavenccai")) = False Then
        fechacai.Value = datPrimaryRS.Recordset.Fields("fechavenccai")
    Else
        fechacai.Value = "00/00/0000"
    End If

If datPrimaryRS.Recordset.EOF = False Then
    If datPrimaryRS.Recordset.Fields("ccocontado") = "S" Then
        Check1(0).Value = 0
        Check1(1).Value = 1
    Else
        Check1(0).Value = 1
        Check1(1).Value = 0
    End If
End If

End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
If datPrimaryRS.Recordset.EOF = False Then

    If IsNull(datPrimaryRS.Recordset.Fields("fechavenccai")) = False Then
        fechacai.Value = datPrimaryRS.Recordset.Fields("fechavenccai")
    Else
        fechacai.Value = "00/00/0000"
    End If

    If datPrimaryRS.Recordset.Fields("ccocontado") = "S" Then
        Check1(0).Value = 0
        Check1(1).Value = 1
    Else
        Check1(0).Value = 1
        Check1(1).Value = 0
    End If
End If
End Sub



Private Sub ec_Click()

    frmproveedores.cuitprov = maskcuit.Text
    frmproveedores.nomprove = txtFields(2).Text
    ecproveedores.Show
    
    
End Sub

Private Sub DataGrid2_Click()

End Sub

Private Sub DataList1_GotFocus()

    If Inicio.opcion1 = True Then
        datcuentas.RecordSource = "select listacuentas.* from listacuentas WHERE empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' ORDER BY IDCUENTA"
        datcuentas.Refresh
        DataList1.ListField = "codigo"
    End If
    If Inicio.opcion2 = True Then
        datcuentas.RecordSource = "select listacuentas.* from listacuentas WHERE empre = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' ORDER BY nombre"
        datcuentas.Refresh
        DataList1.ListField = "nombre"
    End If

End Sub

Private Sub DataList1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        If DataList1.Text <> "" Then Text4.Text = DataList1.BoundText
        DataList1.Visible = False
        grabar.SetFocus
    End If


End Sub

Private Sub DataList1_LostFocus()

    DataList2.Visible = False

End Sub

Private Sub DataList2_GotFocus()

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

End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        If DataList2.Text <> "" Then txtFields(11).Text = DataList2.BoundText
        DataList2.Visible = False
        SSTab1.Tab = 1
        fechacai.SetFocus
    End If

End Sub

Private Sub DataList2_LostFocus()

    DataList2.Visible = False

End Sub

Private Sub fechacai_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Text3.SetFocus
    End If

End Sub

Private Sub Form_GotFocus()
        maskcuit.Mask = ""
        maskcuit.MaxLength = 13
End Sub

Private Sub Form_Load()
frmproveedores.Top = 0
    
    Inicio.Toolbar1.Visible = True


datacontrib.ConnectionString = login.conexiontotal
datbusca.ConnectionString = login.conexiontotal
datcuentas.ConnectionString = login.conexiontotal
datPrimaryRS.ConnectionString = login.conexiontotal

SSTab1.Tab = 0

If login.empresaact1 > 0 Then
    empresareal = login.empresaact1
Else
    empresareal = login.empresaact
End If

    datacontrib.RecordSource = "select condtrib.* from condtrib"
    datacontrib.Refresh

    datcuentas.RecordSource = "select listacuentas.* from listacuentas WHERE empre = " & login.empresaact & " and inicioper = '" & login.iper & "' ORDER BY IDCUENTA"
    datcuentas.Refresh
    datPrimaryRS.RecordSource = "select proveedores.* from proveedores where empresa = " & empresareal & " ORDER BY razonsocial"
    datPrimaryRS.Refresh
    
    DataCombo2.Text = txtFields(2).Text
    maskcuit.Mask = "##-########-#"
    maskcuit.MaxLength = 13
    If datPrimaryRS.Recordset.EOF = True Then
            datPrimaryRS.Recordset.AddNew
            txtFields(1) = empresareal
            maskcuit.SelLength = 13
            maskcuit.SelText = ""
    End If
    
    If IsNull(datPrimaryRS.Recordset.Fields("fechavenccai")) = False Then
         fechacai.Value = datPrimaryRS.Recordset.Fields("fechavenccai")
    End If
    
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
On Error Resume Next
     If Check1(0).Value = 0 And Check1(1).Value = 0 Then
        mensa = MsgBox("Debe ingresar si es Proveedor de Cta.Cte. o Contado", vbInformation, "Atención")
        Check1(0).SetFocus
        Exit Sub
     End If
        
     If IsNull(fechacai.Value) = False And fechacai.Value <> "00/00/0000" Then
        datPrimaryRS.Recordset.Fields("fechavenccai") = fechacai.Value
     End If
        
     If Check1(1).Value = 1 Then
        datPrimaryRS.Recordset.Fields("ccocontado") = "S"
     Else
        datPrimaryRS.Recordset.Fields("ccocontado") = ""
     End If
     datPrimaryRS.Recordset.UpdateBatch adAffectCurrent
     datPrimaryRS.Refresh
     
    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Proveedores"
    Inicio.datauditoria.Recordset.Fields("accion") = "Modificacion Prove.:" + txtFields(2).Text
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent
     
     nuevo.SetFocus
fuera:
End Sub

Private Sub maskcuit_Change()

    If KeyAscii = 13 Then
        KeyAscii = 0
        KeyAscii = 9
    End If

End Sub

Private Sub maskcuit_LostFocus()

        Call busca_Click
        mensa = verifica_cuit(maskcuit.Text)

End Sub

Private Sub nuevo_Click()

    datPrimaryRS.Recordset.AddNew
    txtFields(1) = empresareal
    maskcuit.SelLength = 13
    maskcuit.SelText = ""
    
    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Proveedores"
    Inicio.datauditoria.Recordset.Fields("accion") = "Alta Prove.:"
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent
    
    
    txtFields(2).SetFocus

End Sub

Private Sub ordenarazosocial_Click()

    datPrimaryRS.RecordSource = "select proveedores.* from proveedores  WHERE proveedores.empresa = " & empresareal & " ORDER BY razonsocial"
    datPrimaryRS.Refresh
    
End Sub

Private Sub ordenarcodigo_Click()

    datPrimaryRS.RecordSource = "select proveedores.* from proveedores WHERE proveedores.empresa = " & empresareal & " ORDER BY codproveedor"
    datPrimaryRS.Refresh
    
End Sub

Private Sub salir_Click()

    Unload Me

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)

     If KeyAscii = 13 Then
        KeyAscii = 0
        Text4.SetFocus
     End If

End Sub

Private Sub Text4_GotFocus()

    DataList1.Visible = True
    DataList1.SetFocus

End Sub

Private Sub txtFields_Change(Index As Integer)
On Error GoTo fuera

If datPrimaryRS.Recordset.EOF = False Then
    If datPrimaryRS.Recordset.Fields("ccocontado") = "S" Then
        Check1(0).Value = 0
        Check1(1).Value = 1
    End If
End If

fuera:
End Sub

Private Sub txtFields_GotFocus(Index As Integer)

If Index = 11 Then
    DataList2.Visible = True
    DataList2.SetFocus
End If


End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Index = 2 Then
            DataCombo1.SetFocus
            Exit Sub
        End If
        If Index = 3 Then
            maskcuit.SetFocus
            Exit Sub
        End If

        txtFields(Index + 1).SetFocus
    End If
End Sub

