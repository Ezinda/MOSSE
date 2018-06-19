Attribute VB_Name = "verificacuenta"
Public Function vericuenta(cuenta As Double, cuentainvalida) As Integer
   
    bases.datcuentas.ConnectionString = login.conexiontotal
    bases.datcuentas.RecordSource = "select cuentas.* from cuentas where empre = " & login.empresaact & " and imp = 'S'"
    bases.datcuentas.Refresh
    bases.datcuentas.Recordset.Filter = "codcontable = " & cuenta & ""
    
    If bases.datcuentas.Recordset.EOF = True Then
        cuentainvalida = 1
        MsgBox "No Existe esta cuenta contable", vbCritical, "Verificar"
        Exit Function
    End If
    cuentainvalida = 0

End Function
