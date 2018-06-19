VERSION 5.00
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#62.0#0"; "vbskpro2.ocx"
Begin VB.Form frmrecibocobroasigna 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de Cobros sin Comprobantes"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14520
   Icon            =   "frmrecibocobroasigna.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   14520
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "importe"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   1
      EndProperty
      DataSource      =   "databonan"
      Height          =   285
      Index           =   11
      Left            =   6120
      TabIndex        =   58
      Top             =   4800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton asiglista 
      Caption         =   "Aceptar Asignación de &Lista"
      Height          =   375
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   6000
      UseMaskColor    =   -1  'True
      Width           =   3615
   End
   Begin VB.ListBox List1 
      Height          =   5460
      Left            =   12000
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   55
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton asignartodos 
      Height          =   495
      Left            =   12000
      Picture         =   "frmrecibocobroasigna.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton limpia 
      Caption         =   "Limpia Items"
      Height          =   495
      Left            =   12480
      Picture         =   "frmrecibocobroasigna.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton nuevaorden 
      Caption         =   "&Nueva Asig."
      Height          =   735
      Left            =   8400
      Picture         =   "frmrecibocobroasigna.frx":0DB6
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   840
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton salir 
      Caption         =   "&Cerrar"
      Height          =   735
      Left            =   9720
      Picture         =   "frmrecibocobroasigna.frx":11F8
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   840
      Width           =   1335
   End
   Begin MSDataListLib.DataList DataList1 
      Bindings        =   "frmrecibocobroasigna.frx":163A
      Height          =   2400
      Left            =   480
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   4233
      _Version        =   393216
      MatchEntry      =   -1  'True
      BackColor       =   12648447
      ListField       =   "razonsocial"
      BoundColumn     =   "codcliente"
   End
   Begin MSMask.MaskEdBox totalabonan 
      Height          =   255
      Left            =   2040
      TabIndex        =   47
      Top             =   3120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "$#,##0.00;-$#,##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      DataField       =   "fecha"
      DataSource      =   "datordendepago"
      Height          =   375
      Left            =   5280
      TabIndex        =   38
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   840
      Picture         =   "frmrecibocobroasigna.frx":1657
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   840
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin MSDataListLib.DataList DataList4 
      Bindings        =   "frmrecibocobroasigna.frx":1A99
      Height          =   2400
      Left            =   5280
      TabIndex        =   35
      Top             =   1560
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   3810
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "comp"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataList DataList3 
      Bindings        =   "frmrecibocobroasigna.frx":1AB7
      Height          =   2160
      Left            =   5280
      TabIndex        =   34
      Top             =   1560
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   3810
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "nrorden"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataList DataList2 
      Bindings        =   "frmrecibocobroasigna.frx":1AD2
      Height          =   1230
      Left            =   7440
      TabIndex        =   11
      Top             =   2520
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1693
      _Version        =   393216
      MatchEntry      =   -1  'True
      BackColor       =   12632256
      ListField       =   "comp"
      BoundColumn     =   "id"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton limpiar 
      Caption         =   "&Limpiar"
      Height          =   375
      Left            =   3480
      TabIndex        =   24
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "comprobante"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      DataSource      =   "databonan"
      Height          =   285
      Index           =   10
      Left            =   7440
      TabIndex        =   23
      Top             =   3960
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "comprobante"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      DataSource      =   "databonan"
      Height          =   285
      Index           =   9
      Left            =   7440
      TabIndex        =   22
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "comprobante"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      DataSource      =   "databonan"
      Height          =   285
      Index           =   8
      Left            =   7440
      TabIndex        =   21
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "comprobante"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      DataSource      =   "databonan"
      Height          =   285
      Index           =   7
      Left            =   7440
      TabIndex        =   20
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "comprobante"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      DataSource      =   "databonan"
      Height          =   285
      Index           =   6
      Left            =   7440
      TabIndex        =   19
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "comprobante"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      DataSource      =   "databonan"
      Height          =   285
      Index           =   5
      Left            =   7440
      TabIndex        =   18
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "comprobante"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      DataSource      =   "databonan"
      Height          =   285
      Index           =   4
      Left            =   7440
      TabIndex        =   17
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton aceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1320
      TabIndex        =   16
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "comprobante"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   1
      EndProperty
      DataSource      =   "databonan"
      Height          =   285
      Index           =   3
      Left            =   7440
      TabIndex        =   15
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "fechacompro"
      DataSource      =   "databonan"
      Height          =   285
      Index           =   1
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      DataField       =   "nomcliente"
      DataSource      =   "databonan"
      Height          =   285
      Index           =   0
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2280
      Width           =   3375
   End
   Begin MSMask.MaskEdBox totalfactura 
      Height          =   375
      Left            =   5040
      TabIndex        =   12
      Top             =   3360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   12648447
      Enabled         =   0   'False
      Format          =   "    #,##0.00;(    #,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.PictureBox saldocompro 
      Height          =   255
      Index           =   0
      Left            =   9480
      ScaleHeight     =   195
      ScaleWidth      =   1515
      TabIndex        =   25
      Top             =   2280
      Width           =   1575
   End
   Begin VB.PictureBox saldocompro 
      Height          =   255
      Index           =   1
      Left            =   9480
      ScaleHeight     =   195
      ScaleWidth      =   1515
      TabIndex        =   26
      Top             =   2520
      Width           =   1575
   End
   Begin VB.PictureBox saldocompro 
      Height          =   255
      Index           =   2
      Left            =   9480
      ScaleHeight     =   195
      ScaleWidth      =   1515
      TabIndex        =   27
      Top             =   2760
      Width           =   1575
   End
   Begin VB.PictureBox saldocompro 
      Height          =   255
      Index           =   3
      Left            =   9480
      ScaleHeight     =   195
      ScaleWidth      =   1515
      TabIndex        =   28
      Top             =   3240
      Width           =   1575
   End
   Begin VB.PictureBox saldocompro 
      Height          =   255
      Index           =   4
      Left            =   9480
      ScaleHeight     =   195
      ScaleWidth      =   1515
      TabIndex        =   29
      Top             =   3000
      Width           =   1575
   End
   Begin VB.PictureBox saldocompro 
      Height          =   255
      Index           =   5
      Left            =   9480
      ScaleHeight     =   195
      ScaleWidth      =   1515
      TabIndex        =   30
      Top             =   3480
      Width           =   1575
   End
   Begin VB.PictureBox saldocompro 
      Height          =   255
      Index           =   6
      Left            =   9480
      ScaleHeight     =   195
      ScaleWidth      =   1515
      TabIndex        =   31
      Top             =   3720
      Width           =   1575
   End
   Begin VB.PictureBox saldocompro 
      Height          =   255
      Index           =   7
      Left            =   9480
      ScaleHeight     =   195
      ScaleWidth      =   1515
      TabIndex        =   32
      Top             =   3960
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "frmrecibocobroasigna.frx":1AF0
      Height          =   855
      Left            =   720
      TabIndex        =   33
      Top             =   5280
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1508
      _Version        =   393216
      BackColor       =   12648447
      Enabled         =   -1  'True
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   15
      BeginProperty Column00 
         DataField       =   "fecha"
         Caption         =   "fecha"
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
         DataField       =   "cliente"
         Caption         =   "cliente"
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
         DataField       =   "inicioper"
         Caption         =   "inicioper"
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
         DataField       =   "finper"
         Caption         =   "finper"
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
         DataField       =   "numcompr"
         Caption         =   "numcompr"
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
         DataField       =   "total"
         Caption         =   "total"
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
         DataField       =   "tipocompr"
         Caption         =   "tipocompr"
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
         DataField       =   "comp"
         Caption         =   "comp"
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
         DataField       =   "saldo"
         Caption         =   "saldo"
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
         DataField       =   "cdt"
         Caption         =   "cdt"
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
         DataField       =   "id"
         Caption         =   "id"
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
         DataField       =   "imputado"
         Caption         =   "imputado"
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
      BeginProperty Column12 
         DataField       =   "contado"
         Caption         =   "contado"
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
      BeginProperty Column13 
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
      BeginProperty Column14 
         DataField       =   "codcliente"
         Caption         =   "codcliente"
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
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
         EndProperty
         BeginProperty Column11 
         EndProperty
         BeginProperty Column12 
         EndProperty
         BeginProperty Column13 
         EndProperty
         BeginProperty Column14 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   255
      Left            =   5520
      TabIndex        =   7
      Top             =   4800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   8281
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "&RECIBOS"
      TabPicture(0)   =   "frmrecibocobroasigna.frx":1B0E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Text1(2)"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(2)=   "Command3"
      Tab(0).Control(3)=   "Label9"
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(5)=   "Shape7"
      Tab(0).Control(6)=   "Label1"
      Tab(0).Control(7)=   "Shape1"
      Tab(0).Control(8)=   "Label5"
      Tab(0).Control(9)=   "Label4"
      Tab(0).Control(10)=   "Shape2"
      Tab(0).Control(11)=   "Shape4"
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "&NOTAS DE CREDITOS"
      TabPicture(1)   =   "frmrecibocobroasigna.frx":1B2A
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Shape5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label6"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Shape6"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label7"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Shape3"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Shape8"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label10"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Frame4"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "montonc"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Command2"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).ControlCount=   12
      Begin VB.CommandButton Command2 
         Caption         =   "&Asignar Orden"
         Enabled         =   0   'False
         Height          =   735
         Left            =   7200
         Picture         =   "frmrecibocobroasigna.frx":1B46
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   840
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         DataField       =   "importe"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "##,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         DataSource      =   "databonan"
         Height          =   285
         Index           =   2
         Left            =   -69120
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   2280
         Width           =   1455
      End
      Begin VB.PictureBox montonc 
         Height          =   290
         Left            =   5880
         ScaleHeight     =   225
         ScaleWidth      =   1395
         TabIndex        =   48
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Frame Frame4 
         Caption         =   "Nº de Comprobante"
         Height          =   855
         Left            =   2400
         TabIndex        =   9
         Top             =   720
         Width           =   2655
         Begin MSDataListLib.DataCombo DataCombo2 
            Bindings        =   "frmrecibocobroasigna.frx":1F88
            Height          =   315
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   741
            _Version        =   393216
            MatchEntry      =   -1  'True
            ListField       =   "comp"
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
      End
      Begin VB.Frame Frame1 
         Caption         =   "Nº Orden"
         Height          =   855
         Left            =   -72480
         TabIndex        =   4
         Top             =   720
         Width           =   2415
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "frmrecibocobroasigna.frx":1FA6
            Height          =   360
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   741
            _Version        =   393216
            MatchEntry      =   -1  'True
            ListField       =   "nrorden"
            BoundColumn     =   "inicioper"
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
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Asignar Orden"
         Enabled         =   0   'False
         Height          =   735
         Left            =   -67800
         Picture         =   "frmrecibocobroasigna.frx":1FC1
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   840
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Conceptos que se abonan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   255
         Left            =   600
         TabIndex        =   46
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Shape Shape8 
         Height          =   2655
         Left            =   480
         Top             =   1800
         Width           =   10815
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Comprobante"
         Height          =   255
         Left            =   -69840
         TabIndex        =   45
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Comprobante"
         Height          =   255
         Left            =   5160
         TabIndex        =   44
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo:"
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
         Left            =   -73800
         TabIndex        =   43
         Top             =   3120
         Width           =   735
      End
      Begin VB.Shape Shape7 
         Height          =   495
         Left            =   -73920
         Top             =   3000
         Width           =   3135
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo:"
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
         Left            =   1200
         TabIndex        =   42
         Top             =   3120
         Width           =   735
      End
      Begin VB.Shape Shape3 
         Height          =   495
         Left            =   1080
         Top             =   3000
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Conceptos que se abonan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   255
         Left            =   -74400
         TabIndex        =   6
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Shape Shape1 
         Height          =   2655
         Left            =   -74520
         Top             =   1800
         Width           =   10815
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha de Pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   255
         Left            =   5280
         TabIndex        =   41
         Top             =   600
         Width           =   1575
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H8000000C&
         Height          =   855
         Left            =   5160
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   255
         Left            =   600
         TabIndex        =   40
         Top             =   480
         Width           =   855
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H8000000C&
         Height          =   975
         Left            =   480
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha de Nota"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   255
         Left            =   -69720
         TabIndex        =   39
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   255
         Left            =   -74400
         TabIndex        =   37
         Top             =   480
         Width           =   855
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H8000000C&
         Height          =   975
         Left            =   -74520
         Top             =   600
         Width           =   1815
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H8000000C&
         Height          =   855
         Left            =   -69840
         Top             =   720
         Width           =   1935
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid4 
      Bindings        =   "frmrecibocobroasigna.frx":2403
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      _Version        =   393216
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   75
      BeginProperty Column00 
         DataField       =   "id"
         Caption         =   "id"
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
         DataField       =   "fecha"
         Caption         =   "fecha"
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
         DataField       =   "cliente"
         Caption         =   "cliente"
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
         DataField       =   "tipocompr"
         Caption         =   "tipocompr"
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
         DataField       =   "numcompr"
         Caption         =   "numcompr"
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
         DataField       =   "col1"
         Caption         =   "col1"
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
         DataField       =   "col2"
         Caption         =   "col2"
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
         DataField       =   "col3"
         Caption         =   "col3"
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
         DataField       =   "col4"
         Caption         =   "col4"
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
      BeginProperty Column12 
         DataField       =   "col5"
         Caption         =   "col5"
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
      BeginProperty Column13 
         DataField       =   "col6"
         Caption         =   "col6"
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
      BeginProperty Column14 
         DataField       =   "col7"
         Caption         =   "col7"
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
      BeginProperty Column15 
         DataField       =   "col8"
         Caption         =   "col8"
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
      BeginProperty Column16 
         DataField       =   "col9"
         Caption         =   "col9"
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
      BeginProperty Column17 
         DataField       =   "col10"
         Caption         =   "col10"
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
      BeginProperty Column18 
         DataField       =   "col11"
         Caption         =   "col11"
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
      BeginProperty Column19 
         DataField       =   "col12"
         Caption         =   "col12"
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
      BeginProperty Column20 
         DataField       =   "col13"
         Caption         =   "col13"
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
      BeginProperty Column21 
         DataField       =   "col14"
         Caption         =   "col14"
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
      BeginProperty Column22 
         DataField       =   "col15"
         Caption         =   "col15"
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
      BeginProperty Column23 
         DataField       =   "cuenta"
         Caption         =   "cuenta"
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
      BeginProperty Column24 
         DataField       =   "total"
         Caption         =   "total"
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
      BeginProperty Column25 
         DataField       =   "cerrado"
         Caption         =   "cerrado"
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
      BeginProperty Column26 
         DataField       =   "cd1"
         Caption         =   "cd1"
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
      BeginProperty Column27 
         DataField       =   "ch1"
         Caption         =   "ch1"
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
      BeginProperty Column28 
         DataField       =   "cd2"
         Caption         =   "cd2"
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
      BeginProperty Column29 
         DataField       =   "ch2"
         Caption         =   "ch2"
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
      BeginProperty Column30 
         DataField       =   "cd3"
         Caption         =   "cd3"
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
      BeginProperty Column31 
         DataField       =   "ch3"
         Caption         =   "ch3"
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
      BeginProperty Column32 
         DataField       =   "cd4"
         Caption         =   "cd4"
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
      BeginProperty Column33 
         DataField       =   "ch4"
         Caption         =   "ch4"
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
      BeginProperty Column34 
         DataField       =   "cd5"
         Caption         =   "cd5"
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
      BeginProperty Column35 
         DataField       =   "ch5"
         Caption         =   "ch5"
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
      BeginProperty Column36 
         DataField       =   "cd6"
         Caption         =   "cd6"
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
      BeginProperty Column37 
         DataField       =   "ch6"
         Caption         =   "ch6"
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
      BeginProperty Column38 
         DataField       =   "cd7"
         Caption         =   "cd7"
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
      BeginProperty Column39 
         DataField       =   "ch7"
         Caption         =   "ch7"
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
      BeginProperty Column40 
         DataField       =   "cd8"
         Caption         =   "cd8"
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
      BeginProperty Column41 
         DataField       =   "ch8"
         Caption         =   "ch8"
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
      BeginProperty Column42 
         DataField       =   "cd9"
         Caption         =   "cd9"
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
      BeginProperty Column43 
         DataField       =   "ch9"
         Caption         =   "ch9"
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
      BeginProperty Column44 
         DataField       =   "cd10"
         Caption         =   "cd10"
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
      BeginProperty Column45 
         DataField       =   "ch10"
         Caption         =   "ch10"
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
      BeginProperty Column46 
         DataField       =   "cd11"
         Caption         =   "cd11"
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
      BeginProperty Column47 
         DataField       =   "ch11"
         Caption         =   "ch11"
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
      BeginProperty Column48 
         DataField       =   "cd12"
         Caption         =   "cd12"
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
      BeginProperty Column49 
         DataField       =   "ch12"
         Caption         =   "ch12"
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
      BeginProperty Column50 
         DataField       =   "cd13"
         Caption         =   "cd13"
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
      BeginProperty Column51 
         DataField       =   "ch13"
         Caption         =   "ch13"
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
      BeginProperty Column52 
         DataField       =   "cd14"
         Caption         =   "cd14"
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
      BeginProperty Column53 
         DataField       =   "ch14"
         Caption         =   "ch14"
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
      BeginProperty Column54 
         DataField       =   "cd15"
         Caption         =   "cd15"
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
      BeginProperty Column55 
         DataField       =   "ch15"
         Caption         =   "ch15"
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
      BeginProperty Column56 
         DataField       =   "cdt"
         Caption         =   "cdt"
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
      BeginProperty Column57 
         DataField       =   "cht"
         Caption         =   "cht"
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
      BeginProperty Column58 
         DataField       =   "asentado"
         Caption         =   "asentado"
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
      BeginProperty Column59 
         DataField       =   "asiento"
         Caption         =   "asiento"
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
      BeginProperty Column60 
         DataField       =   "inicioper"
         Caption         =   "inicioper"
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
      BeginProperty Column61 
         DataField       =   "finper"
         Caption         =   "finper"
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
      BeginProperty Column62 
         DataField       =   "ccosto"
         Caption         =   "ccosto"
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
      BeginProperty Column63 
         DataField       =   "contado"
         Caption         =   "contado"
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
      BeginProperty Column64 
         DataField       =   "nombretarjeta"
         Caption         =   "nombretarjeta"
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
      BeginProperty Column65 
         DataField       =   "codoperacion"
         Caption         =   "codoperacion"
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
      BeginProperty Column66 
         DataField       =   "numordenpub"
         Caption         =   "numordenpub"
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
      BeginProperty Column67 
         DataField       =   "avisador"
         Caption         =   "avisador"
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
      BeginProperty Column68 
         DataField       =   "producto"
         Caption         =   "producto"
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
      BeginProperty Column69 
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
      BeginProperty Column70 
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
      BeginProperty Column71 
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
      BeginProperty Column72 
         DataField       =   "numletras"
         Caption         =   "numletras"
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
      BeginProperty Column73 
         DataField       =   "saldo"
         Caption         =   "saldo"
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
      BeginProperty Column74 
         DataField       =   "imputado"
         Caption         =   "imputado"
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
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
         EndProperty
         BeginProperty Column11 
         EndProperty
         BeginProperty Column12 
         EndProperty
         BeginProperty Column13 
         EndProperty
         BeginProperty Column14 
         EndProperty
         BeginProperty Column15 
         EndProperty
         BeginProperty Column16 
         EndProperty
         BeginProperty Column17 
         EndProperty
         BeginProperty Column18 
         EndProperty
         BeginProperty Column19 
         EndProperty
         BeginProperty Column20 
         EndProperty
         BeginProperty Column21 
         EndProperty
         BeginProperty Column22 
         EndProperty
         BeginProperty Column23 
         EndProperty
         BeginProperty Column24 
         EndProperty
         BeginProperty Column25 
         EndProperty
         BeginProperty Column26 
         EndProperty
         BeginProperty Column27 
         EndProperty
         BeginProperty Column28 
         EndProperty
         BeginProperty Column29 
         EndProperty
         BeginProperty Column30 
         EndProperty
         BeginProperty Column31 
         EndProperty
         BeginProperty Column32 
         EndProperty
         BeginProperty Column33 
         EndProperty
         BeginProperty Column34 
         EndProperty
         BeginProperty Column35 
         EndProperty
         BeginProperty Column36 
         EndProperty
         BeginProperty Column37 
         EndProperty
         BeginProperty Column38 
         EndProperty
         BeginProperty Column39 
         EndProperty
         BeginProperty Column40 
         EndProperty
         BeginProperty Column41 
         EndProperty
         BeginProperty Column42 
         EndProperty
         BeginProperty Column43 
         EndProperty
         BeginProperty Column44 
         EndProperty
         BeginProperty Column45 
         EndProperty
         BeginProperty Column46 
         EndProperty
         BeginProperty Column47 
         EndProperty
         BeginProperty Column48 
         EndProperty
         BeginProperty Column49 
         EndProperty
         BeginProperty Column50 
         EndProperty
         BeginProperty Column51 
         EndProperty
         BeginProperty Column52 
         EndProperty
         BeginProperty Column53 
         EndProperty
         BeginProperty Column54 
         EndProperty
         BeginProperty Column55 
         EndProperty
         BeginProperty Column56 
         EndProperty
         BeginProperty Column57 
         EndProperty
         BeginProperty Column58 
         EndProperty
         BeginProperty Column59 
         EndProperty
         BeginProperty Column60 
         EndProperty
         BeginProperty Column61 
         EndProperty
         BeginProperty Column62 
         EndProperty
         BeginProperty Column63 
         EndProperty
         BeginProperty Column64 
         EndProperty
         BeginProperty Column65 
         EndProperty
         BeginProperty Column66 
         EndProperty
         BeginProperty Column67 
         EndProperty
         BeginProperty Column68 
         EndProperty
         BeginProperty Column69 
         EndProperty
         BeginProperty Column70 
         EndProperty
         BeginProperty Column71 
         EndProperty
         BeginProperty Column72 
         EndProperty
         BeginProperty Column73 
         EndProperty
         BeginProperty Column74 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cancelar 
      Cancel          =   -1  'True
      Caption         =   "Command4"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc datordendepago 
      Height          =   330
      Left            =   480
      Top             =   5760
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
   Begin MSAdodcLib.Adodc databonan 
      Height          =   330
      Left            =   2640
      Top             =   5640
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
      Connect         =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBString     =   "Provider=MSDASQL.1;Password=25072004;Persist Security Info=True;User ID=lucva;Data Source=contable;Initial Catalog=contablesql"
      OLEDBFile       =   ""
      DataSourceName  =   "contable"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select recibocobroabonan.* from recibocobroabonan"
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
   Begin MSAdodcLib.Adodc datinstrumento 
      Height          =   330
      Left            =   3840
      Top             =   5640
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
   Begin MSAdodcLib.Adodc datproveedores 
      Height          =   330
      Left            =   4920
      Top             =   5640
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
   Begin MSAdodcLib.Adodc datconsultacomp 
      Height          =   330
      Left            =   6240
      Top             =   5640
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
   Begin MSAdodcLib.Adodc datlibrocompras 
      Height          =   330
      Left            =   6000
      Top             =   5520
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
   Begin MSAdodcLib.Adodc datcuentas 
      Height          =   330
      Left            =   5640
      Top             =   5400
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
      Left            =   720
      Top             =   5760
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
   Begin MSAdodcLib.Adodc datasiento 
      Height          =   330
      Left            =   1920
      Top             =   5760
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
   Begin MSRDC.MSRDC reporte 
      Height          =   375
      Left            =   8760
      Top             =   5880
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
   Begin MSAdodcLib.Adodc datordensinc 
      Height          =   330
      Left            =   360
      Top             =   6000
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
   Begin Crystal.CrystalReport CrystalReporte 
      Left            =   6720
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowTitle     =   "Orden de Pago"
   End
   Begin MSAdodcLib.Adodc datnotascredito 
      Height          =   330
      Left            =   5640
      Top             =   6000
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
   Begin vbskpro.Skinner Skinner1 
      Left            =   120
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
      LcK2            =   $"frmrecibocobroasigna.frx":2421
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
   Begin VB.PictureBox impfactura 
      Height          =   255
      Left            =   10200
      ScaleHeight     =   195
      ScaleWidth      =   1515
      TabIndex        =   57
      Top             =   5160
      Width           =   1575
   End
   Begin VB.PictureBox saldolista 
      Height          =   255
      Left            =   10200
      ScaleHeight     =   195
      ScaleWidth      =   1515
      TabIndex        =   59
      Top             =   5520
      Width           =   1575
   End
   Begin VB.PictureBox totalorden 
      Height          =   255
      Left            =   10200
      ScaleHeight     =   195
      ScaleWidth      =   1515
      TabIndex        =   60
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Orden:"
      Height          =   255
      Left            =   8160
      TabIndex        =   63
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Comp.Seleccionados:"
      Height          =   255
      Left            =   8160
      TabIndex        =   62
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo de Orden ref.a lista:"
      Height          =   255
      Left            =   8160
      TabIndex        =   61
      Top             =   5520
      Width           =   1935
   End
End
Attribute VB_Name = "frmrecibocobroasigna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim importeapagar As Double
Dim totalab As Currency
Dim totalinst(50) As Currency
Dim detalleint(50) As String
Dim totalconc(50) As Currency
Dim nrocompro(50) As String
Dim cuentaint(50) As Integer
Dim nomprov As String
Dim saldoactual As Currency
Dim saldo As Currency
Dim Cuenta As Integer
Dim codprove As Double
Dim idlibrogrid(50) As Integer
Dim saldolibro(50) As Currency
Dim sincomp As Integer
Dim codigopago As Integer
Dim importeorden(10) As Currency
Public numorden As String
Dim empresareal As Integer
Dim totalfac As Currency
Dim inicioperiodo As Date
Dim listaasigna As Integer


Private Sub borrar_Click()
On Error GoTo erroreliminar

    nrocompro(DataGrid1.Row) = ""
    databonan.Recordset.Delete adAffectCurrent
    databonan.Refresh
Exit Sub

erroreliminar:
MsgBox "No se pudo eliminar Concepto"
    
End Sub


Private Sub aceptar_Click()
On Error GoTo fueraerror

    If saldolista.Value <> totalorden.Value Then saldo = saldolista.Value
    If totalfac >= saldo Then
        saldoanterior = saldo
        saldo = 0
    Else
       If saldolista.Value > 0 And saldolista.Value < totalorden.Value Then saldo = saldolista.Value
       saldo = saldo - totalfac
    End If
        
    totalabonan.Text = saldo
    If saldo > 0 Then
        saldocompro(Cuenta - 3).Value = 0
        importeorden(Cuenta - 3) = totalfac
    End If
    If saldo < 0 Then
        saldocompro(Cuenta - 3).Value = totalfac - saldoanterior
        importeorden(Cuenta - 3) = saldoanterior
    End If
    If saldo = 0 Then
        saldocompro(Cuenta - 3).Value = totalfac - saldoanterior
        importeorden(Cuenta - 3) = saldoanterior
    End If
    
  
If saldo = 0 Or (saldo > 0 And saldo < 0.1) Or (saldo < 0 And saldo > -0.1) Then
    For x = 3 To 10
        Text1(x).Enabled = False
    Next x
    Command3.Enabled = True
    command2.Enabled = True
    aceptar.Enabled = False
    Command3.SetFocus
    Exit Sub
End If
    totalfac = 0
    Text1(Cuenta + 1).SetFocus
Exit Sub
fueraerror:
    mensa = MsgBox("Demaciadas facturas para asignar", vbCritical, "Error")


End Sub

Private Sub asiglista_Click()


If listaasigna = 0 Then databonan.Recordset.Delete adAffectCurrent

For x = 0 To List1.ListCount - 1
List1.ListIndex = x
If List1.Selected(x) = True Then
        databonan.Recordset.AddNew
        databonan.Recordset.Fields("nrorden") = DataCombo1.Text
        databonan.Recordset.Fields("empresa") = login.empresaact
        databonan.Recordset.Fields("inicioper") = login.iper
        databonan.Recordset.Fields("finper") = login.fper
        databonan.Recordset.Fields("comprobante") = List1.Text
        databonan.Recordset.Fields("nomcliente") = nomprov
        databonan.Recordset.Fields("codcliente") = codprove
        tipocomprob = Left(List1.Text, 3)
        comprob = Right(List1.Text, 13)
        If Left(tipocomprob, 1) = " " Then
            tipocomprob = " "
            For Y = 1 To 13
                car = Mid(comprob, Y, 1)
                If car <> " " Then GoTo finne
            Next Y
finne:
            comprob = Right(comprob, 14 - Y)
        End If
        datlibrocompras.RecordSource = "select libroventas.* from libroventas where empresa = " & login.empresaact & " and numcompr = '" & comprob & "' and tipocompr = '" & tipocomprob & "'"
        datlibrocompras.Refresh

        databonan.Recordset.Fields("fechacompro") = datlibrocompras.Recordset.Fields("fecha")
        databonan.Recordset.Fields("importe") = datlibrocompras.Recordset.Fields("total")
    Rem    If IsNull(datlibrocompras.Recordset.Fields("total")) = false Then databonan.Recordset.Fields("importe") = datlibrocompras.Recordset.Fields("saldo")
        
        databonan.Recordset.Fields("saldofactura") = 0
        databonan.Recordset.UpdateBatch adAffectCurrent
        
        datlibrocompras.Recordset.Fields("saldo") = 0
        datlibrocompras.Recordset.Fields("imputado") = "S"
        datlibrocompras.Recordset.UpdateBatch adAffectCurrent
End If
Next x
 Unload Me
 frmrecibocobroasigna.Show


End Sub

Private Sub asignartodos_Click()
i = 0

Do While Not i = List1.ListCount
   
   List1.Selected(i) = True
   i = i + 1

Loop

End Sub

Private Sub Command1_Click()


    DataList1.Visible = True
    If SSTab1.Tab = 0 Then
        DataList3.Visible = True
    Else
        DataList4.Visible = True
    End If
    DataList1.SetFocus
    


End Sub



Private Sub Command2_Click()
Dim compro(10) As String

    For x = 3 To 10
             compro(x) = Text1(x).Text
    Next x
        
    If Text1(3).Text = "" Then GoTo fin
    codpro = codprove
    nompro = Text1(0).Text
    fechap = MaskEdBox1.Text
    comprob = Right(compro(3), 13)
    tipocomprob = Left(compro(3), 4)
    If tipocomprob = "   " Then
        tipocomprob = " "
        For x = 1 To 13
            car = Mid(compro(3), x, 1)
            If car <> " " Then GoTo finne
        Next x
finne:
        comprob = Right(compro(3), 16 - x)
    End If
    databonan.Recordset.AddNew
    databonan.Recordset.Fields(0) = DataCombo2.Text
    databonan.Recordset.Fields("empresa") = login.empresaact
    databonan.Recordset.Fields("inicioper") = login.iper
    databonan.Recordset.Fields("finper") = login.fper
    databonan.Recordset.Fields(7) = Text1(3).Text
    databonan.Recordset.Fields("importe") = importeorden(0)
    databonan.Recordset.Fields("saldofactura") = Val(saldocompro(0).Value)
    databonan.Recordset.Fields("comprobante") = compro(3)
    databonan.Recordset.Fields("nomcliente") = nompro
    databonan.Recordset.Fields("codcliente") = codpro
    databonan.Recordset.Fields("fechacompro") = fechap
    databonan.Recordset.UpdateBatch adAffectCurrent
        
    datlibrocompras.RecordSource = "select libroventas.* from libroventas where empresa = " & login.empresaact & " and numcompr = '" & comprob & "' and tipocompr = '" & tipocomprob & "'"
    datlibrocompras.Refresh
    datlibrocompras.Recordset.Fields("saldo") = Val(saldocompro(0).Value)
    datlibrocompras.Recordset.Fields("imputado") = "S"
    datlibrocompras.Recordset.UpdateBatch adAffectCurrent
    
    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Asignación de Notas de Credito"
    Inicio.datauditoria.Recordset.Fields("accion") = "Asignación Nota de Credito " + DataCombo2.Text
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent

    For x = 4 To 10
        If compro(x) = "" Then GoTo fin
        databonan.Recordset.AddNew
        databonan.Recordset.Fields("nrorden") = DataCombo2.Text
        databonan.Recordset.Fields("empresa") = login.empresaact
        databonan.Recordset.Fields("inicioper") = login.iper
        databonan.Recordset.Fields("finper") = login.fper
        databonan.Recordset.Fields("comprobante") = compro(x)
        databonan.Recordset.Fields("nomcliente") = nompro
        databonan.Recordset.Fields("codcliente") = codpro
        databonan.Recordset.Fields("fechacompro") = fechap
        databonan.Recordset.Fields("importe") = importeorden(x - 3)
        databonan.Recordset.Fields("saldofactura") = Val(saldocompro(x - 3).Value)
        databonan.Recordset.UpdateBatch adAffectCurrent
        comprob = Right(compro(x), 13)
        tipocomprob = Left(compro(x), 4)
        
                datlibrocompras.RecordSource = "select libroventas.* from libroventas where empresa = " & login.empresaact & " and numcompr = '" & comprob & "' and tipocompr = '" & tipocomprob & "'"
                datlibrocompras.Refresh
                datlibrocompras.Recordset.Fields("saldo") = Val(saldocompro(x - 3).Value)
                datlibrocompras.Recordset.Fields("imputado") = "S"
                datlibrocompras.Recordset.UpdateBatch adAffectCurrent
   Next x
    
fin:
    
    comprob = Right(DataCombo2.Text, 13)
    tipocomprob = Left(DataCombo2.Text, 4)
 

    datlibrocompras.RecordSource = "select libroventas.* from libroventas where empresa = " & login.empresaact & " and numcompr = '" & comprob & "' and tipocompr = '" & tipocomprob & "'"
    datlibrocompras.Refresh
    datlibrocompras.Recordset.Fields("saldo") = 0
    datlibrocompras.Recordset.Fields("imputado") = "S"
    datlibrocompras.Recordset.UpdateBatch adAffectCurrent



DataList2.Visible = False
Call nuevaorden_Click




End Sub

Private Sub Command3_Click()
Dim compro(10) As String

    For x = 3 To 10
             compro(x) = Text1(x).Text
    Next x
    
    If Text1(3).Text = "" Then GoTo fin
    databonan.Recordset.Fields(7) = Text1(3).Text
    databonan.Recordset.Fields("importe") = importeorden(0)
    databonan.Recordset.Fields("saldofactura") = Val(saldocompro(0).Value)
    databonan.Recordset.UpdateBatch adAffectCurrent
    codpro = databonan.Recordset.Fields("codcliente")
    nompro = databonan.Recordset.Fields("nomcliente")
    fechap = databonan.Recordset.Fields("fechacompro")
    tipocomprob = Left(compro(3), 3)
    comprob = Right(compro(3), 13)
    If tipocomprob = "   " Then
        tipocomprob = " "
        For x = 1 To 13
            car = Mid(compro(3), x, 1)
            If car <> " " Then GoTo finne
        Next x
finne:
        comprob = Right(compro(3), 16 - x)
    End If
    datlibrocompras.RecordSource = "select libroventas.* from libroventas where empresa = " & login.empresaact & " and numcompr = '" & comprob & "' and tipocompr = '" & tipocomprob & "'"
    datlibrocompras.Refresh
    datlibrocompras.Recordset.Fields("saldo") = Val(saldocompro(0).Value)
    datlibrocompras.Recordset.Fields("imputado") = "S"
    datlibrocompras.Recordset.UpdateBatch adAffectCurrent
    
    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Asignación de Cobros sin Comprobantes"
    Inicio.datauditoria.Recordset.Fields("accion") = "Asignación Recibo: " + DataCombo1.Text
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent

    For x = 4 To 10
        If compro(x) = "" Then GoTo fin
        databonan.Recordset.AddNew
        databonan.Recordset.Fields("nrorden") = DataCombo1.Text
        databonan.Recordset.Fields("empresa") = login.empresaact
        databonan.Recordset.Fields("inicioper") = login.iper
        databonan.Recordset.Fields("finper") = login.fper
        databonan.Recordset.Fields("comprobante") = compro(x)
        databonan.Recordset.Fields("nomcliente") = nompro
        databonan.Recordset.Fields("codcliente") = codpro
        databonan.Recordset.Fields("fechacompro") = fechap
        databonan.Recordset.Fields("importe") = importeorden(x - 3)
        databonan.Recordset.Fields("saldofactura") = Val(saldocompro(x - 3).Value)
        databonan.Recordset.UpdateBatch adAffectCurrent
        tipocomprob = Left(compro(x), 3)
        comprob = Right(compro(x), 13)
                datlibrocompras.RecordSource = "select libroventas.* from libroventas where empresa = " & login.empresaact & " and numcompr = '" & comprob & "' and tipocompr = '" & tipocomprob & "'"
                datlibrocompras.Refresh
                datlibrocompras.Recordset.Fields("saldo") = Val(saldocompro(x - 3).Value)
                datlibrocompras.Recordset.Fields("imputado") = "S"
                datlibrocompras.Recordset.UpdateBatch adAffectCurrent
   Next x

fin:
    
    If Val(saldolista.Value) >= 0 And Val(saldolista.Value) < Val(totalorden.Value) Then
                listaasigna = 1
                asiglista.Enabled = True
                asiglista.SetFocus
                SendKeys "{ENTER}", True
                Exit Sub
    End If


Unload Me
frmrecibocobroasigna.Show



End Sub

Private Sub Command4_Click()
Dim tabla As String
Dim tabla1 As String
Dim ruta As String

ruta = "\Empresa" + Right(Str(login.empresaact), Len(Str(login.empresaact)) - 1)

Rem reporte.SQL = "consultaordesnpago.nrorden, consultaordesnpago.empresa, consultaordesnpago.nomproveedor, consultaordesnpago.comprobante, consultaordesnpago.fechacompro, consultaordesnpago.importe, consultaordesnpago.id, consultaordesnpago.razonsocial, consultaordesnpago.cuit, consultaordesnpago.domicilio, consultaordesnpago.localidad, consultaordesnpago.fecha, consultaordesnpago.domprov, consultaordesnpago.locprov, consultaordesnpago.cuitprov, consultaordesnpago.saldofactura FROM contablesql.dbo.consultaordesnpago consultaordesnpago WHERE consultaordesnpago.nrorden= '" & frmordendepago.numorden & "' and consultaordesnpago.empresa = " & login.empresaact & " ORDER BY consultaordesnpago.razonsocial ASC, consultaordesnpago.id ASC"
Rem tabla = reporte.SQL

With CrystalReporte
    .ReportFileName = App.Path & ruta + "\Ordendepago.rpt"
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
 Rem   .Destination = crptToWindow
    .Destination = crptToPrinter
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .Action = 1
    .ReportFileName = App.Path & ruta + "\Ordendepago1.rpt"
    .Connect = login.conexionreporte
    .DiscardSavedData = True
    .RetrieveDataFiles
    .ReportSource = 0
    .SQLQuery = tabla
 Rem   .Destination = crptToWindow
    .Destination = crptToPrinter
    .PrintFileType = crptCrystal
    .WindowState = crptMaximized
    .Action = 1
      
End With
End Sub

Private Sub Command5_Click()

    DataGrid1.Columns(6).Caption = "Detalle de Pago"
    DataGrid1.Columns(7).Visible = False
    DataGrid1.Columns(8).Caption = "Fecha"
    DataGrid1.Columns(10).Visible = True
    DataGrid1.Columns(10).Width = 1395
    DataGrid1.Columns(11).Locked = False
    sincomp = 1

End Sub


Private Sub DataCombo1_KeyPress(KeyAscii As Integer)
On Error GoTo errormod
    If KeyAscii = 13 Then
        KeyAscii = 0
        If DataCombo1.Text = "" Then
          Rem   Command1.SetFocus
            Exit Sub
        End If
        If Right(DataCombo1.Text, 4) <> "-LCM" Then
            DataCombo1.Text = Mid("00000000", 1, 8 - Len(DataCombo1.Text)) + DataCombo1.Text
        End If
        Command3.Enabled = False
        command2.Enabled = False
        
        inicioperiodo = DataCombo1.BoundText
        databonan.RecordSource = "select recibocobroabonan.* from recibocobroabonan WHERE recibocobroabonan.empresa = " & login.empresaact & " and inicioper = '" & inicioperiodo & "' and nrorden = '" & DataCombo1.Text & "' Order by fechacompro, id"
        databonan.Refresh
        datinstrumento.RecordSource = "select recibocobroinstrumento.* from recibocobroinstrumento WHERE recibocobroinstrumento.empresa = " & login.empresaact & " and inicioper = '" & inicioperiodo & "' and nrorden = '" & DataCombo1.Text & "' Order by id"
        datinstrumento.Refresh
        codprove = databonan.Recordset.Fields(5)
        nomprov = databonan.Recordset.Fields(6)
        datconsultacomp.RecordSource = "select conceptoscobran1.* from conceptoscobran1 where empresa = " & login.empresaact & " and codcliente = " & codprove & " order by comp"
        datconsultacomp.Refresh
        datordendepago.RecordSource = "select recibocobro.* from recibocobro WHERE recibocobro.empresa = " & login.empresaact & " and inicioper = '" & inicioperiodo & "' and nrorden = '" & DataCombo1.Text & "' Order by nrorden"
        datordendepago.Refresh
        
        If datconsultacomp.Recordset.EOF = True Then Exit Sub
        List1.Clear
        listaasigna = 0
        importe = 0
        totalorden.Value = Val(Text1(2).Text)
        i = 0
        datconsultacomp.Recordset.MoveFirst
        Do While Not datconsultacomp.Recordset.EOF
   
            listatxt = datconsultacomp.Recordset.Fields("comp")
            List1.AddItem listatxt
            List1.Selected(i) = False
            i = i + 1
            datconsultacomp.Recordset.MoveNext
        Loop
        saldolista.Value = totalorden.Value
        
        
    End If
Rem     databonan.Recordset.MoveLast
    
If IsNull(saldo) = False Then saldo = Text1(2).Text
     Text1(3).SetFocus
errormod:

End Sub


Private Sub DataGrid1_ButtonClick(ByVal ColIndex As Integer)

If ColIndex = 7 Then
        DataList2.Visible = True
        DataList2.Left = DataGrid1.Columns(7).Left + DataGrid1.Left
        DataList2.Width = DataGrid1.Columns(7).Width
        DataList2.Top = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.RowHeight
        DataList2.SetFocus
End If
    
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
On Error GoTo erroringreso

  
    If KeyAscii = 13 And DataGrid1.Col = 7 Then
        If DataGrid1.Columns(7).Text = "" Then
                    DataList2.Visible = True
                    DataList2.Left = DataGrid1.Columns(7).Left + DataGrid1.Left
                    DataList2.Width = DataGrid1.Columns(7).Width
                    DataList2.Top = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.RowHeight
                    KeyAscii = 0
                    DataList2.SetFocus
                    Exit Sub
        Else
            KeyAscii = 9
        End If
    End If
    If KeyAscii = 13 And DataGrid1.Col = 8 Then
          KeyAscii = 9
    End If
    If KeyAscii = 13 And DataGrid1.Col = 9 Then
        KeyAscii = 9
    End If
    
    If KeyAscii = 13 And DataGrid1.Col = 11 And sincomp = 1 Then
        KeyAscii = 0
        nuevo.SetFocus
    End If
Exit Sub
erroringreso:
    Call nuevo.SetFocus



End Sub

Private Sub DataCombo2_KeyPress(KeyAscii As Integer)
On Error GoTo errormod
    If KeyAscii = 13 Then
        KeyAscii = 0
        If DataCombo2.Text = "" Then
          Rem   Command1.SetFocus
            Exit Sub
        End If
        Command3.Enabled = False
        command2.Enabled = False
        datnotascredito.RecordSource = "select conceptoscobran1.* from conceptoscobran1 where empresa = " & login.empresaact & " and comp = '" & DataCombo2.Text & "' order by comp"
        datnotascredito.Refresh

        MaskEdBox1.Text = datnotascredito.Recordset.Fields("fecha")

        codprove = datnotascredito.Recordset.Fields("codcliente")

        datconsultacomp.RecordSource = "select conceptoscobran1.* from conceptoscobran1 where empresa = " & login.empresaact & " and codcliente = " & codprove & " and comp <> 'NCA' and comp <> 'NCB' order by comp"
        datconsultacomp.Refresh
        
        comprob = Right(DataCombo2.Text, 13)
        tipocomprob = Left(DataCombo2.Text, 4)
        
        datlibrocompras.RecordSource = "select libroventas.* from libroventas where empresa = " & login.empresaact & " and numcompr = '" & comprob & "' and tipocompr = '" & tipocomprob & "'"
        datlibrocompras.Refresh
        
        Text1(0).Text = datlibrocompras.Recordset.Fields("cliente")
        Text1(1).Text = datlibrocompras.Recordset.Fields("fecha")
        montonc.Value = Val(datlibrocompras.Recordset.Fields("total")) * -1
        Text1(2).Text = montonc.Value
        
    End If
    datnotascredito.Recordset.MoveLast
    
        
    datnotascredito.RecordSource = "select conceptoscobran1.* from conceptoscobran1 where empresa = " & login.empresaact & " and (tipocompr = 'NCA' or tipocompr = 'NCB') order by comp"
    datnotascredito.Refresh
    
If IsNull(saldo) = False Then saldo = Val(Text1(2).Text)
    Text1(3).SetFocus
errormod:
End Sub

Private Sub DataList1_Click()

    datordensinc.RecordSource = "select consultarecibosinc.* from consultarecibosinc where consultarecibosinc.empresa = " & login.empresaact & " and codcliente = " & DataList1.BoundText & " order by nrorden"
    datordensinc.Refresh

    datnotascredito.RecordSource = "select conceptoscobran1.* from conceptoscobran1 where empresa = " & login.empresaact & " and cliente = '" & DataList1.Text & "' and (tipocompr = 'NCA' or tipocompr = 'NCB') order by comp"
    datnotascredito.Refresh

End Sub



Private Sub DataList1_KeyDown(KeyCode As Integer, Shift As Integer)

    If DataList1.BoundText <> "" Then
        datordensinc.RecordSource = "select consultarecibosinc.* from consultarecibosinc where consultarecibosinc.empresa = " & login.empresaact & " and codcliente = " & DataList1.BoundText & " order by nrorden"
        datordensinc.Refresh
        
        datnotascredito.RecordSource = "select conceptoscobran1.* from conceptoscobran1 where empresa = " & login.empresaact & " and cliente = '" & DataList1.Text & "' and (tipocompr = 'NCA' or tipocompr = 'NCB') order by comp"
        datnotascredito.Refresh
    End If
    
End Sub

Private Sub DataList1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
            If SSTab1.Tab = 0 Then
                     DataList3.SetFocus
            Else
                       DataList4.SetFocus
            End If
    End If
    
End Sub

Private Sub DataList2_GotFocus()

If SSTab1.Tab = 0 Then
   datconsultacomp.RecordSource = "select conceptoscobran1.* from conceptoscobran1 where empresa = " & login.empresaact & " and codcliente = " & codprove & " order by comp"
   datconsultacomp.Refresh
Else
   datconsultacomp.RecordSource = "select conceptoscobran1.* from conceptoscobran1 where empresa = " & login.empresaact & " and codcliente = " & codprove & " and comp <> 'NCA' and comp <> 'NCB' order by comp"
   datconsultacomp.Refresh
End If

End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            Text1(Cuenta).Text = DataList2.Text
            compro = DataList2.Text

            If compro = "" Then Exit Sub
   
            For x = (Cuenta - 1) To 3 Step -1
                If Text1(Cuenta).Text = Text1(x).Text Then
                    mensa = MsgBox("Este comprobante ya fue ingresado, cambielo", vbCritical, "!! Error !!")
                    Text1(Cuenta).SetFocus
                    Exit Sub
                End If
            Next x
            
                For Y = 0 To List1.ListCount - 1
                    List1.ListIndex = Y
                    If List1.Selected(Y) = True Then
                        If Text1(Cuenta).Text = List1.Text Then
                            mensa = MsgBox("Este comprobante ya fue ingresado, cambielo", vbCritical, "!! Error !!")
                            Text1(Cuenta).SetFocus
                            Exit Sub
                        End If
                    End If
                Next Y
                        
            datconsultacomp.RecordSource = "select conceptoscobran1.* from conceptoscobran1 where empresa = " & login.empresaact & " and codcliente = " & codprove & " and comp = '" & compro & "' order by comp"
            datconsultacomp.Refresh
            If datconsultacomp.Recordset.Fields("contado") = "S" Then
                mensa = MsgBox("Esta factura es de Contado, no se puede asignar", vbCritical, "Error")
                Text1(Cuenta).SetFocus
                Exit Sub
            End If
            If DataGrid3.Columns(8).Text = "" Then
                DataGrid3.Columns(5).Visible = True
                DataGrid3.Columns(8).Visible = False
                totalfac = DataGrid3.Columns(5).Text
            Else
                DataGrid3.Columns(5).Visible = False
                DataGrid3.Columns(8).Visible = True
                totalfac = DataGrid3.Columns(8).Text
            End If
            totalfactura.Text = totalfac
            DataList2.Visible = False
            aceptar.SetFocus
    End If
End Sub

Private Sub DataList2_LostFocus()
            
            DataList2.Visible = False
            
End Sub


Private Sub DataList3_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        DataCombo1.Text = DataList3.Text
        DataCombo1.SetFocus
        DataList3.Visible = False
        DataList1.Visible = False
    End If
    

End Sub

Private Sub DataList4_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        DataCombo2.Text = DataList4.Text
        DataCombo2.SetFocus
        DataList4.Visible = False
        DataList1.Visible = False
    End If

End Sub

Private Sub DataList4_LostFocus()

    DataList4.Visible = False
    DataList1.Visible = False


End Sub

Private Sub Form_Load()

    Inicio.Toolbar1.Visible = True


frmrecibocobroasigna.Top = 0
frmrecibocobroasigna.Left = 0

datasiento.ConnectionString = login.conexiontotal
datconsultacomp.ConnectionString = login.conexiontotal
datcuentas.ConnectionString = login.conexiontotal
datinstrumento.ConnectionString = login.conexiontotal
datlibrocompras.ConnectionString = login.conexiontotal
datmaestro.ConnectionString = login.conexiontotal
datordendepago.ConnectionString = login.conexiontotal
datproveedores.ConnectionString = login.conexiontotal
datordensinc.ConnectionString = login.conexiontotal
datnotascredito.ConnectionString = login.conexiontotal

If login.empresaact1 > 0 Then
    empresareal = login.empresaact1
Else
    empresareal = login.empresaact
End If

SSTab1.Tab = 0

  datordendepago.RecordSource = "select recibocobro.* from recibocobro WHERE recibocobro.empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' and nrorden = '0'"
  datordendepago.Refresh

    datordensinc.RecordSource = "select consultarecibosinc.* from consultarecibosinc where consultarecibosinc.empresa = " & login.empresaact & " order by nrorden"
    datordensinc.Refresh
    

    datnotascredito.RecordSource = "select conceptoscobran1.* from conceptoscobran1 where empresa = " & login.empresaact & " and (tipocompr = 'NCA' or tipocompr = 'NCB') order by comp"
    datnotascredito.Refresh
    
 databonan.RecordSource = "select recibocobroabonan.* from recibocobroabonan WHERE recibocobroabonan.empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' and  nrorden = '0'"
 databonan.Refresh
  
 datinstrumento.RecordSource = "select recibocobroinstrumento.* from recibocobroinstrumento WHERE recibocobroinstrumento.empresa = " & login.empresaact & " and inicioper = '" & login.iper & "' and finper = '" & login.fper & "' and nrorden = '" & orden & "' Order by id"
  datinstrumento.Refresh
  
  datproveedores.RecordSource = "select clientes.* from clientes where empresa = " & empresareal & " order by razonsocial"
  datproveedores.Refresh
  

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Inicio.Toolbar1.Visible = False

End Sub

Private Sub limpia_Click()
i = 0

Do While Not i = List1.ListCount
   
   List1.Selected(i) = False
   i = i + 1

Loop
impfactura.Value = 0
saldolista.Value = totalorden.Value
asiglista.Enabled = False
End Sub
Private Sub List1_ItemCheck(Item As Integer)

    If List1.Selected(List1.ListIndex) = True Then
            datconsultacomp.RecordSource = "select conceptoscobran1.* from conceptoscobran1 where empresa = " & login.empresaact & " and codcliente = " & codprove & " and comp = '" & List1.Text & "'"
            datconsultacomp.Refresh
            If DataGrid3.Columns(8).Text = "" Then
                DataGrid3.Columns(5).Visible = True
                DataGrid3.Columns(8).Visible = False
                importe = DataGrid3.Columns(5).Text
            Else
                DataGrid3.Columns(5).Visible = False
                DataGrid3.Columns(8).Visible = True
                importe = DataGrid3.Columns(8).Text
            End If
            impfactura.Value = importe + Val(impfactura.Value)
    End If
    If List1.Selected(List1.ListIndex) = False Then
            datconsultacomp.RecordSource = "select conceptoscobran1.* from conceptoscobran1 where empresa = " & login.empresaact & " and codcliente = " & codprove & " and comp = '" & List1.Text & "'"
            datconsultacomp.Refresh
            If DataGrid3.Columns(8).Text = "" Then
                DataGrid3.Columns(5).Visible = True
                DataGrid3.Columns(8).Visible = False
                importe = DataGrid3.Columns(5).Text
            Else
                DataGrid3.Columns(5).Visible = False
                DataGrid3.Columns(8).Visible = True
                importe = DataGrid3.Columns(8).Text
            End If
            impfactura.Value = Val(impfactura.Value) - importe
    End If

    saldolista.Value = totalorden.Value - impfactura.Value
    If saldolista.Value = 0 Or (saldolista.Value <= 0.09 And saldolista.Value >= -0.09) Then
        asiglista.Enabled = True
        asiglista.SetFocus
    Else
        asiglista.Enabled = False
    End If
    

End Sub

Private Sub limpiar_Click()
                      
    saldo = Text1(2).Text

    For x = 3 To 10
        Text1(x).Text = ""
        saldocompro(x - 3).Value = 0
        Text1(x).Enabled = True
    Next x
    aceptar.Enabled = True
    saldo = Text1(2).Text
    totalabonan = saldo
    totalfactura.Text = ""
    Command3.Enabled = False
    command2.Enabled = False
    Text1(3).SetFocus
    
End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        nuevo.SetFocus
    End If
    
End Sub

Private Sub nuevaorden_Click()
   
DataCombo1.Text = ""
totalabonan.Text = ""
    For x = 3 To 10
        Text1(x).Text = ""
        saldocompro(x - 3).Value = 0
        importeorden(x - 3) = 0
        Text1(x).Enabled = True
    Next x
    aceptar.Enabled = True
    
totalfactura.Text = ""

Call Form_Load



End Sub


Private Sub salir_Click()

    Unload Me

End Sub


Private Sub Text1_GotFocus(Index As Integer)
    
       If Index >= 3 Then
        DataList2.Visible = True
        DataList2.Left = Text1(Index).Left
        DataList2.Width = Text1(Index).Width
        DataList2.Top = Text1(Index).Top + Text1(Index).Height
        Cuenta = Index
        DataList2.SetFocus
       End If
       
End Sub

