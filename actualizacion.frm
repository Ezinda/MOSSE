VERSION 5.00
Begin VB.Form actualizacion 
   Caption         =   "Actualizacion"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0,00 ""$"""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   2
      EndProperty
      Height          =   375
      HideSelection   =   0   'False
      Left            =   1200
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "actualizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Inet1_StateChanged(ByVal State As Integer)

End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_LostFocus()

    

End Sub

Private Sub Form_Load()

End Sub

Private Sub TextBox1_Change()

    TextBox1.maº

End Sub
