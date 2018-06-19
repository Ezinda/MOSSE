VERSION 5.00
Begin VB.Form CONFREGIONAL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración Regional"
   ClientHeight    =   3060
   ClientLeft      =   2505
   ClientTop       =   2955
   ClientWidth     =   6945
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Configuracion Regional"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton salir 
      Caption         =   "&restaurar"
      Height          =   315
      Left            =   3480
      TabIndex        =   25
      Top             =   2520
      Width           =   1035
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   10
      Left            =   5520
      TabIndex        =   23
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   16
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   8
      Left            =   5520
      TabIndex        =   15
      Top             =   1200
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   7
      Left            =   5520
      TabIndex        =   14
      Top             =   960
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   6
      Left            =   5520
      TabIndex        =   13
      Top             =   720
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   5
      Left            =   5520
      TabIndex        =   6
      Top             =   480
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   5
      Top             =   1440
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   4
      Top             =   1200
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   3
      Top             =   960
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   2
      Top             =   720
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   1
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cambiar"
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   2520
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Números"
      Height          =   1575
      Index           =   0
      Left            =   2160
      TabIndex        =   12
      Top             =   240
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Moneda"
      Height          =   1575
      Index           =   1
      Left            =   5160
      TabIndex        =   17
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Fechas: ""dd/MM/aaaa"""
      Height          =   255
      Index           =   10
      Left            =   3360
      TabIndex        =   24
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Horas: ""HH:mm:ss"""
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   22
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Digitos decimales ""2"""
      Height          =   255
      Index           =   8
      Left            =   3240
      TabIndex        =   21
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Separador de miles ""."""
      Height          =   255
      Index           =   7
      Left            =   3240
      TabIndex        =   20
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Punto Decimal "","""
      Height          =   255
      Index           =   6
      Left            =   3240
      TabIndex        =   19
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Simbolo de Moneda ""$"""
      Height          =   255
      Index           =   5
      Left            =   3240
      TabIndex        =   18
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Digitos en grupo ""3"""
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Simbolo Negativo ""-"""
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Digitos decimales ""2"""
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Separador de miles ""."""
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Punto decimal "","""
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "CONFREGIONAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
Dim l As Long
Dim ruta As String
Dim Ret As Long
Dim setea As String
Dim setea1 As String
Dim setea2 As String
Dim setea3 As String

Rem ESTABLECE "." COMO PUNTO DECIMAL
setea = SetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, ".")

Rem ESTABLECE "," COMO separador de miles
setea1 = SetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_STHOUSAND, ",")

Rem ESTABLECE 2 digitos como decimales
If (SetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_IDIGITS, "2")) Then
    Check1(2).Value = 1
Else
    Check1(2).Value = 0
End If
Rem ESTABLECE "-" COMO SIMBOLO NEGATIVO
If (SetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SNEGATIVESIGN, "-")) Then
    Check1(3).Value = 1
Else
    Check1(3).Value = 0
End If
Rem ESTABLECE 3 COMO DIGITOS EN GRUPOS
If (SetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SGROUPING, "3")) Then
    Check1(4).Value = 1
Else
    Check1(4).Value = 0
End If

Rem ESTABLECE "$" COMO SIMBOLO DE MONEDA
If (SetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SCURRENCY, "$")) Then
    Check1(5).Value = 1
Else
    Check1(5).Value = 0
End If
Rem ESTABLECE "." COMO PUNTO DECIMAL DE MONEDA
setea2 = SetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONDECIMALSEP, ".")

Rem ESTABLECE "," COMO SEPARADOR DE MILES EN MONEDA
setea3 = SetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONTHOUSANDSEP, ",")

Rem ESTABLECE "2" decimales EN MONEDA
If (SetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_ICURRDIGITS, "2")) Then
    Check1(8).Value = 1
Else
    Check1(8).Value = 0
End If

Rem ESTABLECE "HH:mm:ss" en horas
If (SetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_STIMEFORMAT, "HH:mm:ss")) Then
    Check1(9).Value = 1
Else
    Check1(9).Value = 0
End If

Rem ESTABLECE "dd/MM/aaaa" en fecha
If (SetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE, "dd/MM/yyyy")) Then
    Check1(10).Value = 1
Else
    Check1(10).Value = 0
End If


Rem ruta = App.Path & "\contable.exe"

Rem Ret = Shell(ruta, vbNormalFocus)
Unload Me


End Sub

Private Sub Form_Load()
Dim i As Integer, aux As String

    
    Call Command1_Click
Rem        Call salir_Click

End Sub


Public Function Base2Long(s As String, ByVal nB As Integer) As Long
Dim s2 As String
Dim i As Long
Dim j As Long
Dim X As Long
Dim n As Boolean
Dim s3 As String

If Len(s) < 1 Then
    Base2Long = 0
    Exit Function
End If

s2 = UCase(s)

If Left$(s2, 1) = "-" Then
    n = True
    s2 = Right$(s2, Len(s2) - 1)
Else
    n = False
End If

j = 1
X = 0

For i = Len(s2) To 1 Step -1
    s3 = Mid$(s2, i, 1)
    Select Case s3
        Case "0" To "9":
            X = X + j * (Asc(s3) - 48)
        Case "A" To "Z":
            X = X + j * (Asc(s3) - 55)
    End Select

    j = j * nB
Next i

If n Then
    X = -X
End If

Base2Long = X
End Function

Private Sub salir_Click()
Dim l As Long
Dim setea As String
Dim setea1 As String
Dim setea2 As String
Dim setea3 As String
Dim ruta As String
Dim Ret As Long

Rem ESTABLECE "," COMO PUNTO DECIMAL
setea = SetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, ",")

Rem ESTABLECE "." COMO separador de miles
setea1 = SetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_STHOUSAND, ".")

Rem ESTABLECE "," COMO PUNTO DECIMAL DE MONEDA
setea2 = SetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONDECIMALSEP, ",")

Rem ESTABLECE "." COMO SEPARADOR DE MILES EN MONEDA
setea3 = SetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONTHOUSANDSEP, ".")

    Unload Me

End Sub
