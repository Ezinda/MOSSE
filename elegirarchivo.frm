VERSION 5.00
Begin VB.Form elegirarchivo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Direccion del Archivo de Pago Facil"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   3720
      Width           =   6615
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2775
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00E0E0E0&
      Height          =   3015
      Left            =   3000
      TabIndex        =   1
      Top             =   600
      Width           =   3855
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00E0E0E0&
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
End
Attribute VB_Name = "elegirarchivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    afip_f_electronica.Text1.Text = Text1.Text
    Unload Me

End Sub

Private Sub Dir1_Change()

    File1.Path = Dir1.Path
    File1.FileName = "*.txt"

    Text1.Text = Dir1.Path
    
    
End Sub

Private Sub Drive1_Change()

    Dir1.Path = Drive1.Drive

End Sub

Private Sub File1_Click()

    Text1.Text = Dir1.Path + "\" + File1.FileName
    
End Sub

Private Sub File1_DblClick()

    Call Command1_Click

End Sub

Private Sub Form_Load()
Aplicar_skin Me
elegirarchivo.Top = 0
elegirarchivo.Left = 0


    Dir1.Path = "c:"

End Sub
