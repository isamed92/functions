Function Det1erVTO(linea As Integer, FechaPago As Date)

'************************************************************************************************
'******** Determina Segun Linea de credito y fecha de pago la fecha 1er Vencimiento   ***********
'************************************************************************************************
On Error GoTo Err_DetFech

Dim RstLinea As Recordset
Dim StrQrylinea As String
     
     FechaLiq = "#" & Trim(Format(FechaPago, "mm/dd/yyyy")) & "#"
          
     StrQrylinea = "SELECT LineasCreditos.LineaCreditoID, LineasCreditos.tipovtoid " & _
                   "FROM LineasCreditos " & _
                   "WHERE (((LineasCreditos.LineaCreditoID)=" & Trim(linea) & "))"

                 
    Set RstLinea = CreditosDB.OpenRecordset(StrQrylinea, dbOpenDynaset, dbSeeChanges)
       
       Dim AuxFecha As Date
       Dim Factor As Integer
       
       Select Case RstLinea!TipoVtoId
              Case 1 ' Mes siguiente, mismo día
                   Det1erVTO = DateAdd("m", 1 + Gracia, FechaPago)
                   
              Case 2 ' 1º día hábil mes siguiente
                   AuxFecha = str(DateAdd("m", 1 + Gracia, FechaPago))
                   AuxFecha = DateSerial(Format(AuxFecha, "yyyy"), Format(AuxFecha, "mm"), "01")
                   Det1erVTO = AuxFecha
                   
              Case 3 '1º día hábil mes sub siguiente
                   AuxFecha = str(DateAdd("m", 2 + Gracia, FechaPago))
                   AuxFecha = DateSerial(Format(AuxFecha, "yyyy"), Format(AuxFecha, "mm"), "01")
                   Det1erVTO = AuxFecha
                   
              Case 5 'Condicional
                   If Day(FechaPago) < 20 Then
                      Factor = 1
                   Else
                      Factor = 2
                   End If
                   AuxFecha = str(DateAdd("m", Factor + Gracia, FechaPago))
                   AuxFecha = DateSerial(Format(AuxFecha, "yyyy"), Format(AuxFecha, "mm"), "01")
                   Det1erVTO = AuxFecha
                   
              Case 6 ' 10 del mes siguiente
                   AuxFecha = str(DateAdd("m", 1 + Gracia, FechaPago))
                   AuxFecha = DateSerial(Format(AuxFecha, "yyyy"), Format(AuxFecha, "mm"), "10")
                   Det1erVTO = AuxFecha
                   
              Case 7 ' 10 del mes subsiguiente
                   AuxFecha = str(DateAdd("m", 2 + Gracia, FechaPago))
                   AuxFecha = DateSerial(Format(AuxFecha, "yyyy"), Format(AuxFecha, "mm"), "10")
                   Det1erVTO = AuxFecha
              
              Case 8 ' ANSES
                   AuxFecha = DateSerial(Year(FechaPago), Month(FechaPago) + 3, 1)
                   AuxFecha = DateAdd("d", -1, AuxFecha)
                   Det1erVTO = AuxFecha
              Case 9 ' ANSES
                   AuxFecha = str(DateAdd("m", 2 + Gracia, FechaPago))
                   AuxFecha = DateSerial(Format(AuxFecha, "yyyy"), Format(AuxFecha, "mm"), "01") - 2
                   Det1erVTO = AuxFecha
       End Select
       
       'Controlo que la fecha calculada no sea un día inhabil
       'Det1erVTO = DameProximoDiaHabil(CDate(Det1erVTO))

    RstLinea.Close
    
 
Exit Function
Err_DetFech:
    MsgBox Err.description, vbCritical, "Atención!!!"
Exit Function

End Function
