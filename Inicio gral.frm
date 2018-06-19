VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.MDIForm Iniciogral 
   AutoShowChildren=   0   'False
   BackColor       =   &H00800000&
   Caption         =   "LUCVA Gestion V.1.0"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10695
   Icon            =   "Inicio gral.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picBackdrop 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1275
      ScaleWidth      =   10635
      TabIndex        =   6
      Top             =   630
      Visible         =   0   'False
      Width           =   10695
      Begin VB.PictureBox picOriginal 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   12900
         Left            =   120
         Picture         =   "Inicio gral.frx":014A
         ScaleHeight     =   12900
         ScaleWidth      =   11550
         TabIndex        =   8
         Top             =   120
         Width           =   11550
      End
      Begin VB.PictureBox picStretched 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   7260
         Left            =   1920
         ScaleHeight     =   7260
         ScaleWidth      =   4095
         TabIndex        =   7
         Top             =   600
         Width           =   4095
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   5
         Left            =   120
         TabIndex        =   16
         Text            =   "Compras Preg.Imprime"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox Check3 
         Height          =   255
         Left            =   10080
         TabIndex        =   14
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   4
         Left            =   8280
         TabIndex        =   13
         Text            =   "Preview Comprobantes"
         Top             =   0
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   3
         Left            =   6840
         TabIndex        =   12
         Text            =   "Multiempresa"
         Top             =   0
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Height          =   255
         Left            =   7920
         TabIndex        =   11
         Top             =   0
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Left            =   6480
         TabIndex        =   10
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   4680
         TabIndex        =   9
         Text            =   "Nº Auto.Libro Ventas"
         Top             =   0
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         Height          =   255
         Left            =   4320
         TabIndex        =   5
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   2280
         TabIndex        =   4
         Text            =   "Ordena Listas por Nombre"
         Top             =   0
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Height          =   255
         Left            =   1920
         TabIndex        =   3
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Text            =   "Ordena Listas por Codigo"
         Top             =   0
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Por Nombre"
         Height          =   315
         Left            =   1800
         Picture         =   "Inicio gral.frx":742BF
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6720
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc datauditoria 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      Top             =   1965
      Visible         =   0   'False
      Width           =   10695
      _ExtentX        =   18865
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
   Begin MSAdodcLib.Adodc datparamgral 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      Top             =   2295
      Visible         =   0   'False
      Width           =   10695
      _ExtentX        =   18865
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
   Begin VB.Menu archivo 
      Caption         =   "&ARCHIVO"
      Begin VB.Menu plandecuentas 
         Caption         =   "Plan de Cuentas"
      End
      Begin VB.Menu libros 
         Caption         =   "Libros"
         WindowList      =   -1  'True
         Begin VB.Menu ivacompras 
            Caption         =   "IVA COMPRAS"
            Begin VB.Menu cargalibrocompras 
               Caption         =   "Carga (Compras Cta.Cte)"
            End
            Begin VB.Menu otrosgastos 
               Caption         =   "Carga (Compras Contado)"
            End
            Begin VB.Menu cierrelibrocompras 
               Caption         =   "Cerrar libro"
            End
            Begin VB.Menu listadolibrocompras 
               Caption         =   "Listado"
            End
            Begin VB.Menu consufactcomp 
               Caption         =   "Consulta Facturas de Compras"
            End
         End
         Begin VB.Menu ivaventas 
            Caption         =   "IVA VENTAS"
            Begin VB.Menu cargalilbroventas 
               Caption         =   "Carga"
            End
            Begin VB.Menu cierrelibroventas 
               Caption         =   "Cerrar libro"
            End
            Begin VB.Menu listadolibroventas 
               Caption         =   "Listado"
            End
            Begin VB.Menu consufacventas 
               Caption         =   "Consulta Facturas de Ventas"
            End
            Begin VB.Menu menuimporta 
               Caption         =   "Importar Libro"
            End
         End
      End
      Begin VB.Menu mincontables 
         Caption         =   "Minutas Contables"
         Begin VB.Menu cargaasientos 
            Caption         =   "Carga"
         End
         Begin VB.Menu verasientos 
            Caption         =   "Ver y/o Modificar"
         End
      End
      Begin VB.Menu menucheques 
         Caption         =   "Cheques"
         Begin VB.Menu menuchequesencartera 
            Caption         =   "En Cartera"
         End
         Begin VB.Menu menuchequesemitidos 
            Caption         =   "Emitidos"
         End
      End
      Begin VB.Menu menuafip 
         Caption         =   "AFIP"
         Begin VB.Menu citiventas 
            Caption         =   "C.I.T.I. Ventas"
         End
      End
      Begin VB.Menu menuimpresoras 
         Caption         =   "Configurar Impresoras"
      End
   End
   Begin VB.Menu empresas 
      Caption         =   "&EMPRESAS"
      Begin VB.Menu confempresa 
         Caption         =   "Datos de Empresas"
      End
      Begin VB.Menu cambianperiodo 
         Caption         =   "Cambiar Periodo de Trabajo"
      End
   End
   Begin VB.Menu proveedores 
      Caption         =   "&PROVEEDORES"
      Begin VB.Menu archproveed 
         Caption         =   "Archivo Proveedores"
      End
      Begin VB.Menu estcuentas 
         Caption         =   "Estados de Cuenta"
      End
      Begin VB.Menu depuracionprov 
         Caption         =   "Depuracion por cambio de Razon Social"
      End
      Begin VB.Menu ordenesdepago 
         Caption         =   "Ordendes de Pago"
         Begin VB.Menu ordpago 
            Caption         =   "Carga"
         End
         Begin VB.Menu asigcomp 
            Caption         =   "Asignar comprobante"
         End
         Begin VB.Menu consord 
            Caption         =   "Consulta"
         End
         Begin VB.Menu listordenes 
            Caption         =   "Listado ordenes emitidas"
         End
         Begin VB.Menu anulorden 
            Caption         =   "Anular Orden de Pago"
         End
      End
   End
   Begin VB.Menu arclientes 
      Caption         =   "&CLIENTES"
      Begin VB.Menu archclientes 
         Caption         =   "Archivo Clientes"
      End
      Begin VB.Menu ecclientes 
         Caption         =   "Estados de Cuenta"
      End
      Begin VB.Menu depuracioncli 
         Caption         =   "Depuracion por cambio de Razon Social"
      End
      Begin VB.Menu facturacionclientes 
         Caption         =   "Facturacion"
         Begin VB.Menu emitirfactura 
            Caption         =   "Emitir Factura"
         End
         Begin VB.Menu consfacturas 
            Caption         =   "Consulta Facturas"
         End
         Begin VB.Menu productos 
            Caption         =   "Productos u Articulos"
         End
      End
      Begin VB.Menu recibos 
         Caption         =   "Recibos"
         Begin VB.Menu emitirrecibo 
            Caption         =   "Emitir Recibo"
         End
         Begin VB.Menu consultarecibos 
            Caption         =   "Consulta Recibos Emitidos"
         End
         Begin VB.Menu listadorecibos 
            Caption         =   "Listado de Recibos Emitidos"
         End
         Begin VB.Menu anurecibo 
            Caption         =   "Anular Recibo"
         End
      End
      Begin VB.Menu asigrecibo 
         Caption         =   "Asignar Comprobante (Recibos y Notas de Credito)"
      End
   End
   Begin VB.Menu menureportes 
      Caption         =   "REPOR&TES"
      Begin VB.Menu librodiario 
         Caption         =   "Libro Diario"
      End
      Begin VB.Menu mayoranalitico 
         Caption         =   "Mayor Analitico"
      End
      Begin VB.Menu sumasysaldos 
         Caption         =   "Balance de Sumas y Saldos"
      End
      Begin VB.Menu sumasysaldosconcc 
         Caption         =   "Balance de Sumas y Saldos Con Centros de Costo"
      End
      Begin VB.Menu informeresul 
         Caption         =   "Informe de Resultados"
      End
   End
   Begin VB.Menu parametros 
      Caption         =   "PA&RAMETROS"
      Begin VB.Menu usuarios 
         Caption         =   "Usuarios"
      End
      Begin VB.Menu plantipo 
         Caption         =   "Importar Configuracion de otro ejercicio"
      End
      Begin VB.Menu importaempresa 
         Caption         =   "Importar Parametrizacion de Otra Empresa"
      End
      Begin VB.Menu importaplan 
         Caption         =   "Importar Plan de Cuentas en Excel"
      End
      Begin VB.Menu paramlibros 
         Caption         =   "Libros"
         Begin VB.Menu conflibro 
            Caption         =   "Configuración de Libro Compras"
         End
         Begin VB.Menu conflibroventas 
            Caption         =   "Configuración de Libro Ventas"
         End
      End
      Begin VB.Menu paramempresa 
         Caption         =   "Empresa"
         Begin VB.Menu digitos 
            Caption         =   "Digitos del codigo contable"
         End
      End
      Begin VB.Menu ventas 
         Caption         =   "Ventas"
         Begin VB.Menu paramfac 
            Caption         =   "Facturacion"
         End
      End
      Begin VB.Menu otparam 
         Caption         =   "Otros Parametros"
      End
      Begin VB.Menu centrocosto 
         Caption         =   "Centros de Costo"
      End
      Begin VB.Menu menuauditoria 
         Caption         =   "Auditoria de Sistema"
      End
      Begin VB.Menu confremoto 
         Caption         =   "Acceso Remoto"
      End
   End
   Begin VB.Menu MODULOS 
      Caption         =   "&MODULOS"
      Begin VB.Menu prgordenes 
         Caption         =   "Ordenes de Publicidad"
      End
      Begin VB.Menu prgfactura 
         Caption         =   "Facturacion"
      End
      Begin VB.Menu sueldos 
         Caption         =   "Sueldos y Jornales"
      End
      Begin VB.Menu acremoto 
         Caption         =   "Permitir Acceso Remoto"
      End
   End
   Begin VB.Menu loginear 
      Caption         =   "&Login"
   End
   Begin VB.Menu acercade 
      Caption         =   "Acerca &de"
      Begin VB.Menu acerca 
         Caption         =   "Acerca de"
      End
      Begin VB.Menu actualizar 
         Caption         =   "Actualizar desde la WEB"
      End
   End
   Begin VB.Menu salir 
      Caption         =   "SALIR"
   End
End
Attribute VB_Name = "Iniciogral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public opcion1 As Boolean
Public opcion2 As Boolean
Public salida As Integer
Public montoefectivo As Currency

Private Sub acremoto_Click()
On Error GoTo fuera
Dim ruta As String
Dim Ret As Long

ruta = App.Path & "\lucvareg.exe"

Ret = Shell(ruta, vbNormalFocus)

    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Acceso Remoto"
    Inicio.datauditoria.Recordset.Fields("accion") = "Registro de IP Publica"
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent

fuera:
End Sub

Private Sub citiventas_Click()

    afipcitiventas.Show

End Sub

Private Sub confremoto_Click()

    frmremoto.Show

End Sub

Private Sub consufactcomp_Click()

    frmconsutalibrocompras.Show

End Sub

Private Sub consufacventas_Click()

    frmconsutalibroventas.Show

End Sub

Private Sub depuracioncli_Click()

    frmdepclientes.Show

End Sub

Private Sub depuracionprov_Click()

    frmdepproveedores.Show

End Sub

Private Sub importaempresa_Click()

    importacuentaotraempresa.Show

End Sub

Private Sub importaplan_Click()

    frmimportaplancuenta.Show

End Sub

Private Sub informeresul_Click()

    frmresultados.Show

End Sub

Private Sub MDIForm_Activate()
    MDIForm_Resize
    
    
    datauditoria.ConnectionString = login.conexiontotal
    datparamgral.ConnectionString = login.conexiontotal
    
    datparamgral.RecordSource = "select parametrosgenerales.* from parametrosgenerales"
    datparamgral.Refresh
    
    datauditoria.RecordSource = "select auditoria.* from auditoria"
    datauditoria.Refresh
    
    datauditoria.Recordset.AddNew
    datauditoria.Recordset.Fields("fecha") = Date
    datauditoria.Recordset.Fields("hora") = Str(Time)
    datauditoria.Recordset.Fields("ventana") = "Login"
    datauditoria.Recordset.Fields("accion") = "Sesión Iniciada"
    datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    datauditoria.Recordset.Fields("empresa") = login.empresaact
    datauditoria.Recordset.UpdateBatch adAffectCurrent
    salida = 1
    
    
    If datparamgral.Recordset.EOF = True Then
        datparamgral.Recordset.AddNew
        datparamgral.Recordset.Fields("montomaxefectivo") = 0
        datparamgral.Recordset.Fields("preguntacai") = 0
        datparamgral.Recordset.Fields("activamultiempresa") = 1
        datparamgral.Recordset.Fields("ordenlista") = 0
        datparamgral.Recordset.Fields("numeroautoventas") = 0
        datparamgral.Recordset.UpdateBatch adAffectCurrent
    End If
    
montoefectivo = datparamgral.Recordset.Fields("montomaxefectivo")
 
 If datparamgral.Recordset.Fields("ordenlista") = 0 Then
    Option1.Value = True
 Else
    Option2.Value = True
 End If
       
If datparamgral.Recordset.Fields("numeroautoventas") < 0 Then
    Check1.Value = datparamgral.Recordset.Fields("numeroautoventas") * -1
Else
    Check1.Value = datparamgral.Recordset.Fields("numeroautoventas")
End If

If datparamgral.Recordset.Fields("activamultiempresa") < 0 Then
    Check2.Value = datparamgral.Recordset.Fields("activamultiempresa") * -1
Else
    Check2.Value = datparamgral.Recordset.Fields("activamultiempresa")
End If

    Check3.Value = 1
  
    
End Sub

' Make the image fit the MDI form.
Private Sub MDIForm_Resize()
    picStretched.Move 0, 0, _
        ScaleWidth, ScaleHeight

    ' Copy the original picture into picStretched.
    picStretched.PaintPicture _
        picOriginal.Picture, _
        0, 0, _
        picStretched.ScaleWidth, _
        picStretched.ScaleHeight, _
        0, 0, _
        picOriginal.ScaleWidth, _
        picOriginal.ScaleHeight
        
    ' Set the MDI form's picture.
    Picture = picStretched.Image
End Sub
Private Sub acerca_Click()
        frmAbout.Show
End Sub

Private Sub actualizar_Click()
Dim ruta As String
Dim Ret As Long

    ruta = App.Path & "\actualizador.exe"

    Ret = Shell(ruta, vbNormalFocus)
    Unload Me
End Sub


Private Sub anulorden_Click()

    frmordenanula.Show

End Sub

Private Sub anurecibo_Click()

    frmreciboanula.Show

End Sub

Private Sub archclientes_Click()

    frmclientes.Show

End Sub

Private Sub archproveed_Click()

    frmproveedores.Show

End Sub

Private Sub asientos_Click()

End Sub

Private Sub asigcomp_Click()

    frmordendepagoasigna.Show

End Sub

Private Sub asigrecibo_Click()

    frmrecibocobroasigna.Show

End Sub

Private Sub cambianperiodo_Click()

    frminicioperiodo.Show

End Sub

Private Sub cargaasientos_Click()

   frmasientos.Show

End Sub

Private Sub cargalibrocompras_Click()

    frmlibrocompras.Show

End Sub

Private Sub cargalilbroventas_Click()

    frmlibroventas.Show

End Sub

Private Sub centrocosto_Click()

    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Centros de Costo"
    Inicio.datauditoria.Recordset.Fields("accion") = "Ingreso a config. de Centros de Costos"
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent


    frmccostos.Show

End Sub

Private Sub cierrelibrocompras_Click()
    
    frmclcompras.Show

End Sub

Private Sub cierrelibroventas_Click()

    frmclventas.Show

End Sub


Private Sub confempresa_Click()
    
    frmEMPRESA.Show

End Sub

Private Sub conflibro_Click()

    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "COLUMNAS COMPRAS"
    Inicio.datauditoria.Recordset.Fields("accion") = "Ingreso a config. Libro Compras"
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent

    frmcolumnascompra.Show

End Sub

Private Sub conflibroventas_Click()

    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "COLUMNAS VENTAS"
    Inicio.datauditoria.Recordset.Fields("accion") = "Ingreso a config. Libro Ventas"
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent

    frmcolumnasventa.Show

End Sub

Private Sub consfacturas_Click()

    frmfacturaconsulta.Show

End Sub

Private Sub consord_Click()

    frmordendepagoconsulta.Show

End Sub

Private Sub consultarecibos_Click()

    frmreciboconsulta.Show

End Sub

Private Sub digitos_Click()
     
    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Configuración de Niveles"
    Inicio.datauditoria.Recordset.Fields("accion") = "Ingreso a config. de niveles"
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent
     
    frmniveles.Show
    
End Sub

Private Sub login_Click()

    login.Show

End Sub

Private Sub ecproveedores_Click()

   

End Sub


Private Sub ecclientes1_Click()

End Sub

Private Sub ecclientes_Click()

    impecclientes.Show

End Sub

Private Sub emitirfactura_Click()
    frmfacclientesradio.Show
End Sub

Private Sub emitirrecibo_Click()

    frmrecibos.Show

End Sub

Private Sub estcuentas_Click()

    impecproved.Show

End Sub

Private Sub facclientesboton_Click()

    frmfacclientes.Show

End Sub

Private Sub importalibro_Click()

    importalibroventas.Show

End Sub

Private Sub librodiario_Click()

    implibrodiario.Show

End Sub

Private Sub listadolibrocompras_Click()
    
    implibrocompras.Show

End Sub

Private Sub listadolibroventas_Click()

    implibroventas.Show
    
End Sub

Private Sub listadorecibos_Click()

    imprecibolistado.Show

End Sub

Private Sub listordenes_Click()

    impordeneslistado.Show

End Sub

Private Sub loginear_Click()
     login.Show
End Sub

Private Sub mayoranalitico_Click()

    impmayoranalitico.Show

End Sub

Private Sub MDIForm_Load()

    Option1.Value = True
    opcion1 = True
    opcion2 = False
    Check1.Value = 0
    Check2.Value = 1

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

        Cancel = salida
        frmsalir.Show

End Sub

Private Sub menuauditoria_Click()

    frmauditoria.Show

End Sub

Private Sub menuchequesemitidos_Click()

    frmcarteraemitidos.Show
    

End Sub

Private Sub menuchequesencartera_Click()
    
    frmcartera.Show

End Sub

Private Sub menuimporta_Click()

    importalibroventas.Show

End Sub

Private Sub menuimpresoras_Click()

    CommonDialog1.ShowPrinter
    CommonDialog1.PrinterDefault = True

End Sub

Private Sub mnuprueba_Click()

    Form1.Show

End Sub

Private Sub Option1_Click()
    If Option1.Value = True Then
        opcion1 = True
        opcion2 = False
    End If
End Sub

Private Sub Option2_Click()

    
    If Option2.Value = True Then
        opcion2 = True
        opcion1 = False
    End If


End Sub

Private Sub ordpago_Click()

    frmordendepago1.Show

End Sub

Private Sub otparam_Click()

    frmotrosparam.Show

End Sub

Private Sub otrosgastos_Click()

    frmotrosgastos.Show

End Sub

Private Sub paramfac_Click()

    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Configuración Parametros de Ventas"
    Inicio.datauditoria.Recordset.Fields("accion") = "Ingreso a config. de parametros de ventas"
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent


    frmparamventas.Show

End Sub

Private Sub plandecuentas_Click()
    

    frmCuentas.Show
    
    
End Sub

Private Sub plantipo_Click()

 importacuenta.Show

End Sub

Private Sub prgfactura_Click()
On Error GoTo fuera
Dim ruta As String
Dim Ret As Long

ruta = App.Path & "\facturacion.exe"

Ret = Shell(ruta, vbNormalFocus)

    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Ordenes de Publicidad"
    Inicio.datauditoria.Recordset.Fields("accion") = "Ingreso a Modulo Facturacion"
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent

    SendKeys login.usuarioactivo, True
    SendKeys "{ENTER}", True
    SendKeys login.contraseña, True
    SendKeys "{ENTER}", True
    SendKeys "{ENTER}", True

fuera:
End Sub

Private Sub prgordenes_Click()
On Error GoTo fuera
Dim ruta As String
Dim Ret As Long

ruta = App.Path & "\ordenes.exe"

Ret = Shell(ruta, vbNormalFocus)

    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Ordenes de Publicidad"
    Inicio.datauditoria.Recordset.Fields("accion") = "Ingreso a Ordenes de Publicidad"
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent

    SendKeys login.usuarioactivo, False
    SendKeys "{ENTER}", False
    SendKeys login.contraseña, False
    SendKeys "{ENTER}", False
    SendKeys "{ENTER}", False

fuera:
End Sub

Private Sub productos_Click()

    frmarticulos.Show

End Sub

Private Sub salir_Click()

    
    frmsalir.Show

End Sub

Private Sub sueldos_Click()
On Error GoTo fuera
Dim ruta As String
Dim Ret As Long

ruta = App.Path & "\sueldos.exe"

Ret = Shell(ruta, vbNormalFocus)

    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "Ordenes de Publicidad"
    Inicio.datauditoria.Recordset.Fields("accion") = "Ingreso a Modulo Facturacion"
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent

    SendKeys login.usuarioactivo, False
    SendKeys "{ENTER}", False
    SendKeys login.contraseña, False
    SendKeys "{ENTER}", False
    SendKeys "{ENTER}", False

fuera:
End Sub

Private Sub sumasysaldos_Click()

    impsumasysaldos.Show

End Sub

Private Sub sumasysaldosconcc_Click()

    impsumasysaldoscc.Show

End Sub

Private Sub usuarios_Click()

    Inicio.datauditoria.ConnectionString = login.conexiontotal
    
    Inicio.datauditoria.RecordSource = "select auditoria.* from auditoria"
    Inicio.datauditoria.Refresh
    
    Inicio.datauditoria.Recordset.AddNew
    Inicio.datauditoria.Recordset.Fields("fecha") = Date
    Inicio.datauditoria.Recordset.Fields("hora") = Str(Time)
    Inicio.datauditoria.Recordset.Fields("ventana") = "USUARIOS"
    Inicio.datauditoria.Recordset.Fields("accion") = "Ingreso a Usuarios"
    Inicio.datauditoria.Recordset.Fields("usuario") = login.usuarioactivo
    Inicio.datauditoria.Recordset.Fields("empresa") = login.empresaact
    Inicio.datauditoria.Recordset.UpdateBatch adAffectCurrent

    frmusuarios.Show

End Sub

Private Sub verasientos_Click()

    frmasientosbusca.Show

End Sub
