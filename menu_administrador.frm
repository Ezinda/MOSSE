VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form menu_administrador 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8550
   ClientLeft      =   2310
   ClientTop       =   1575
   ClientWidth     =   4035
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "menu_administrador.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   4035
   Begin VB.CommandButton acciones 
      Caption         =   "acciones"
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   6120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "menu_administrador.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "menu_administrador.frx":F465
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "menu_administrador.frx":11C17
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "menu_administrador.frx":12069
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "menu_administrador.frx":1956B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "menu_administrador.frx":1FDCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "menu_administrador.frx":272CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "menu_administrador.frx":3A7C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "menu_administrador.frx":4525E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "menu_administrador.frx":4BAC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "menu_administrador.frx":63E97
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "menu_administrador.frx":6A6F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "menu_administrador.frx":77999
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   8535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   15055
      _Version        =   393217
      Indentation     =   88
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "menu_administrador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub acciones_Click()
On Error Resume Next

        xpadre = TreeView1.SelectedItem.Parent
        xindice = TreeView1.SelectedItem
        
''' Nota de Ventas
        If xpadre = "Nota de Ventas" Then
            If xindice = "Nueva" Then
                tipodeventa = 0
                frmnota_venta.Show
            End If
        End If
        
''' Factura de Cta.Cte.
        If xpadre = "Facturacion Cta.Cte" Then
            If xindice = "En Base a Remito" Then
                menu = 1
                frmremitosconsulta.Show
            End If
        End If
        
         If xpadre = "Facturacion Cta.Cte" Then
            If xindice = "En Base a Nota Venta" Then
                menu = 1
                frmcentrodefacturacion.Show
            End If
        End If
        
''' Notas de Credito
        If xpadre = "Notas de Crédito" Then
            If xindice = "Nueva" Then
                frmnota_credito.Show
            End If
        End If
        
        If xpadre = "Notas de Crédito" Then
            If xindice = "Buscar" Then
                lista_notacredito_todas.Show
            End If
        End If
        
''' Notas de Debito
        If xpadre = "Notas de Débito" Then
            If xindice = "Nueva" Then
                frmnota_debito.Show
            End If
        End If
        
        If xpadre = "Notas de Débito" Then
            If xindice = "Buscar" Then
                lista_debitos_todas.Show
            End If
        End If
''' Consulta de Ventas
        If xpadre = "Consultas" Then
            If xindice = "Consultas" Then
                menu = 1
                frmremitosconsulta.Show
            End If
        End If

        If xpadre = "Facturacion Cta.Cte" Then
            If xindice = "Buscar" Then
                lista_facturas_todas.Show
            End If
        End If


''' Cotizaciones
        If xpadre = "Cotizaciones" Then
            If xindice = "Nueva" Then
                menu = 0
                frmpresupuesto.Show
            End If
        End If
        
        If xpadre = "Cotizaciones" Then
            If xindice = "Buscar" Then
                menu = 2
                lista_presupuestos_todos.Show
            End If
        End If


''' Comparativas
        If xpadre = "Comparativas" Then
            If xindice = "Nueva" Then
                menu = 5
                frmcomparativa.Show
            End If
        End If
        
        If xpadre = "Comparativas" Then
            If xindice = "Buscar" Then
                menu = 6
                lista_presupuestos.Show
            End If
        End If


''' Inventarios
        If xpadre = "Gestion Inventario" Then
            If xindice = "Armado de Pedidos" Then
                lista_notadeventas.Show
            End If
        End If

        If xpadre = "Gestion Inventario" Then
            If xindice = "Emision de Remitos" Then
                lista_pendientesremitir.Show
            End If
        End If
        
        If xpadre = "Gestion Inventario" Then
            If xindice = "Remitos Consulta" Then
                menu = 0
                frmremitosconsulta_remito.Show
            End If
        End If

End Sub

Private Sub Form_Load()

Dim hSysmenu As Long
hSysmenu = GetSystemMenu(Me.hWnd, 0)
RemoveMenu hSysmenu, 6, &H400&


'Aplicar_skin Me


menu_administrador.Top = 0
menu_administrador.Left = 0

menu_administrador.Height = Inicio.Height - 1600
TreeView1.Height = menu_administrador.Height - 50


''' Gestion de Ventas
Set nodx = TreeView1.Nodes.Add(, , "MenGestVentas", "Gestion de Ventas")
nodx.Image = 2
Set nodx = TreeView1.Nodes.Add("MenGestVentas", tvwChild, "MenNotaVenta", "Nota de Ventas")
nodx.Image = 7
Set nodx = TreeView1.Nodes.Add("MenNotaVenta", tvwChild, "MenCrearNuevaNV", "Nueva")
nodx.Image = 5

Set nodx = TreeView1.Nodes.Add("MenGestVentas", tvwChild, "MenFacturacion", "Facturacion Cta.Cte")
nodx.Image = 8
Set nodx = TreeView1.Nodes.Add("MenFacturacion", tvwChild, "MenCrearNuevabaseremito", "En Base a Remito")
nodx.Image = 5
Set nodx = TreeView1.Nodes.Add("MenFacturacion", tvwChild, "MenCrearNuevabaseNV", "En Base a Nota Venta")
nodx.Image = 5
Set nodx = TreeView1.Nodes.Add("MenFacturacion", tvwChild, "MenBuscarFac", "Buscar")
nodx.Image = 6

Set nodx = TreeView1.Nodes.Add("MenGestVentas", tvwChild, "MenNotasCredito", "Notas de Crédito")
nodx.Image = 9
Set nodx = TreeView1.Nodes.Add("MenNotasCredito", tvwChild, "MenCrearNuevaNC", "Nueva")
nodx.Image = 5
Set nodx = TreeView1.Nodes.Add("MenNotasCredito", tvwChild, "MenBuscarNC", "Buscar")
nodx.Image = 6

Set nodx = TreeView1.Nodes.Add("MenGestVentas", tvwChild, "MenNotasDebito", "Notas de Débito")
nodx.Image = 10
Set nodx = TreeView1.Nodes.Add("MenNotasDebito", tvwChild, "MenCrearNuevaND", "Nueva")
nodx.Image = 5
Set nodx = TreeView1.Nodes.Add("MenNotasDebito", tvwChild, "MenBuscarND", "Buscar")
nodx.Image = 6



''' Gestion Comercial
Set nodx = TreeView1.Nodes.Add(, , "MenGestComercial", "Gestion Comercial")
nodx.Image = 1
Set nodx = TreeView1.Nodes.Add("MenGestComercial", tvwChild, "MenCotizaciones", "Cotizaciones")
nodx.Image = 3
Set nodx = TreeView1.Nodes.Add("MenCotizaciones", tvwChild, "MenCrearNueva", "Nueva")
nodx.Image = 5
Set nodx = TreeView1.Nodes.Add("MenCotizaciones", tvwChild, "MenBuscar", "Buscar")
nodx.Image = 6

Set nodx = TreeView1.Nodes.Add("MenGestComercial", tvwChild, "MenComparativas", "Comparativas")
nodx.Image = 4
Set nodx = TreeView1.Nodes.Add("MenComparativas", tvwChild, "MenCrearNuevacomp", "Nueva")
nodx.Image = 5
Set nodx = TreeView1.Nodes.Add("MenComparativas", tvwChild, "MenBuscarcomp", "Buscar")
nodx.Image = 6

''' Gestion Inventario
Set nodx = TreeView1.Nodes.Add(, , "MenGestInentario", "Gestion Inventario")
nodx.Image = 11
Set nodx = TreeView1.Nodes.Add("MenGestInentario", tvwChild, "MenArmadoPedidos", "Armado de Pedidos")
nodx.Image = 12
Set nodx = TreeView1.Nodes.Add("MenGestInentario", tvwChild, "MenEmisionRemitos", "Emision de Remitos")
nodx.Image = 13
Set nodx = TreeView1.Nodes.Add("MenGestInentario", tvwChild, "MenRemitosConsulta", "Remitos Consulta")
nodx.Image = 6


End Sub

Private Sub TreeView1_DblClick()

       Call acciones_Click


End Sub

Private Sub TreeView1_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        KeyAscii = 0
        Call acciones_Click
        
    End If

End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next


    

End Sub


