Attribute VB_Name = "Digitoverificador"
Public Function verifica_cuit(cuit As String, cuitinvalido) As String
Dim coeficiente(1 To 10) As Integer
Dim i, sumador, veri_nro, resultado As Integer
Dim cuit_rearmado As String
coeficiente(1) = 5
coeficiente(2) = 4
coeficiente(3) = 3
coeficiente(4) = 2
coeficiente(5) = 7
coeficiente(6) = 6
coeficiente(7) = 5
coeficiente(8) = 4
coeficiente(9) = 3
coeficiente(10) = 2
cuit = Trim(cuit)
cuit_rearmado = ""
For i = 1 To Len(cuit)      'separo cualquier caracter que no tenga que ver con numeros
     If Asc(Mid(cuit, i, 1)) >= 48 And Asc(Mid(cuit, i, 1)) <= 57 Then
        cuit_rearmado = cuit_rearmado & Mid(cuit, i, 1)
     End If
Next
cuit_rearmado = Trim(cuit_rearmado)
If Len(cuit_rearmado) <> 11 Then            ' si to estan todos los digitos
   MsgBox "No estan todos los digitos. ", vbDefaultButton1, "Error en el C.U.I.T."
   cuitinvalido = 1
Else
   sumador = 0
   verificador = Val(Mid(cuit_rearmado, 11, 1)) 'tomo el digito verificador
   For i = 1 To 10
       sumador = sumador + Val(Mid(cuit_rearmado, i, 1)) * coeficiente(i)
       'separo cada digito y lo multiplico por el coeficiente
   Next
   resultado = sumador Mod 11
   resultado = 11 - resultado  'saco el digito verificador
   veri_nro = Val(verificador)
   If veri_nro <> resultado Then
  Rem    MsgBox "No coincide el digito verificador. " & Str(verificador), vbDefaultButton1, "Error en el C.U.I.T."
         MsgBox "CUIT INVALIDO", vbDefaultButton1, "Error en el C.U.I.T."
         cuitinvalido = 1
   Else
      cuit_rearmado = Mid(cuit_rearmado, 1, 2) & "-" & Mid(cuit_rearmado, 3, 8) & "-" & Mid(cuit_rearmado, 11, 1)
      cuitinvalido = 0
   End If
End If
verifica_cuit = cuit_rearmado
End Function
