VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form impresos1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión"
   ClientHeight    =   9570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11745
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   11745
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   9495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "impresos1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

impresos.Top = 0
impresos.Left = 0

End Sub

Private Sub Form_Resize()

    CRViewer1.Width = impresos.Width - 100
    CRViewer1.Height = impresos.Height - 500

End Sub
