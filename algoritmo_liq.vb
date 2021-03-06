Public Sub of202_nanni(LineaId As Integer, Capital As Currency, plazo As Integer, difDias1eraCuota)

On Error GoTo ErrFrances

Dim RstLinea As Recordset
Set RstLinea = CreditosDB.OpenRecordset("SELECT LineasCreditos.* FROM LineasCreditos " & _
                                        "WHERE (((LineasCreditos.LineaCreditoID)=" & LineaId & "))", dbOpenDynaset, dbSeeChanges)

Dim Impuesto As Double
Dim Vartasa As Single
Dim CapitalPrestado, CapitalSolicitado, CapitalPrestadoAjustado As Currency
Dim Seguro, AjustePrimerCuota As Currency
Dim SaldoCapital As Currency
Dim Tempo As Recordset
Dim VarConceptoCuota As Single
Dim VarSeguro As Single
Dim VarFondoGarantia  As Single
Dim DiasAjuste As Integer
Dim datoscargos As String
        
If Not RstLinea.EOF Then
    
    Impuesto = 1 + (RstLinea!LineaIVAGeneral / 100)
    IVASobreInteres = 1 + (RstLinea!LineaIVASobreInt / 100)
    Vartasa = DameTasa(LineaId, plazo)
    CapitalSolicitado = Capital
    DiasAjuste = difDias1eraCuota
    
       'Calculo Ajuste 1º cuota
       If difDias1eraCuota > 0 Then
             AjustePrimerCuota = ((Vartasa / 360) * (difDias1eraCuota-1) * CapitalSolicitado) / 100
       End If

        VarSeguro = RstLinea!LineaSeguroVida
        VarFondoGarantia = RstLinea!LineaFdoGarantia
        If Forms![002AltaCreditos]!AniosFin < 70 Then
           VarFondoGarantia = 0
        Else
           VarSeguro = 0
        End If
        
        If RstLinea!LineaSellado > 0 Then
           SelladoII = RstLinea!LineaSellado / plazo
        End If
        If RstLinea!LineaSeguroVida > 0 Then
           SeguroII = RstLinea!LineaSeguroVida / plazo
        End If
                
        DoCmd.RunSQL ("DELETE Liquidacion.* FROM TempoLiquidacion")
        Set Tempo = CurrentDb.OpenRecordset("TempoLiquidacion")
        
        Dim I As Integer
        Dim fechacuota As Date
        
        CapitalPrestado = Capital
        SaldoCapital = CapitalPrestado
        Tasax = Vartasa / 12 / 100
        Capital = 0
        Interes = 0
        fechacuota = Forms![002AltaCreditos]!PrimerVencimiento
                
        CuotaConIva = Pmt(Tasax * IVASobreInteres, plazo, -CapitalPrestado, 0, 0)
        'Dim SaldoCapital As Long
        CuotaSinIva = Pmt(Tasax, plazo, -CapitalPrestado, 0, 0)
        SaldoCapital = CapitalPrestado
        For I = 1 To plazo
        
            Interes = (SaldoCapital * Tasax)
            InteresConIVA = Redondeo2(Interes * IVASobreInteres)
            Capital = (CuotaSinIva - Interes)
            CuotaPura = Redondeo2(Capital + Interes)
            VarSeguroVida = Redondeo2(CuotaPura * RstLinea!LineaSeguroVida / 100)
            VarFondoGtia = Redondeo2((SaldoCapital / 1000) * VarFondoGarantia)
            SaldoCapital = (SaldoCapital - Capital)
            'Grabo en la tempo
            Tempo.AddNew
            Tempo!Ncuota = I
            Tempo!fechavto = DameProximoDiaHabil(fechacuota)
            Tempo!Tasa = Vartasa
            Tempo!TasaNominalAnual = Vartasa
            Tempo!Solicitado = CapitalPrestado
            Tempo!Prestado = CapitalPrestado
            Tempo!Capital = Capital
            Tempo!Interes = Interes
            Tempo!IvaInteres = Interes * RstLinea!LineaIVASobreInt / 100
            Tempo!SdoCap = SaldoCapital
            Tempo!SeguroVida = 0
            Tempo!FondoGarantia = VarFondoGtia
            Tempo!Gastos = (Capital + Interes) * (RstLinea!LineaGastosAdmin / 100)
            Tempo!IVAGastos = Tempo!Gastos * IVAGeneral / 100
            Tempo!Cuota = Redondeo2(Capital + Tempo!Interes + Tempo!IvaInteres + Tempo!Gastos + Tempo!IVAGastos)
            Tempo!linea = Forms![002AltaCreditos]!ListaLineasHabilitadas.Column(1)
'            Forms![002Altacreditos]!ListaLineasHabilitadas.Column(1), _
'                                Forms![002Altacreditos]![002subaltacreditossolicitante]!NombreII, _
'                                Forms![002Altacreditos]![002subaltacreditossolicitante]!NumeroDoc,
            If Not IsMissing(Forms![002AltaCreditos]![002SubAltaCreditosSolicitante]!NombreII) Then
               Tempo!Cliente = Forms![002AltaCreditos]![002SubAltaCreditosSolicitante]!NombreII
               Tempo!Documento = Forms![002AltaCreditos]![002SubAltaCreditosSolicitante]!NumeroDoc
            End If
            Tempo!Sellado = SelladoII
            Tempo.Update
            fechacuota = DateAdd("m", I * RstLinea!PeriodoAmortizacionID, Forms![002AltaCreditos]!PrimerVencimiento)
        Next
        
        'Calculo Totales Otorgamiento - Fondo Credito
        'Forms![002Altacreditos].Retenciones = ArmoFondoCred(Capital, RstLinea!LineaGastoOriginacion, Forms![002Altacreditos]!AniosFin, DiasAjuste, Vartasa, RstLinea!LineaSellado, _
                                                            RstLinea!LineaGastosAdmin, RstLinea!LineaIVAGeneral, RstLinea!LineaIVAAjuste, 0, Plazo, _
                                                            Capital, 0, RstLinea!destinolineaid, RstLinea!LineaFondoReserva, RstLinea!LineaEdadFondoGtia)

       
      'Borro las retenciones del crédito
       Dim RstRetenciones As Recordset
       Dim Retenciones As Currency
       Retenciones = 0
       DoCmd.RunSQL "DELETE * " & _
                    "FROM RetencionesTempo " & _
                    "WHERE RetencionesTempo.PcIP = '" & Trim(Forms!Menu.IPPc) & "'"

       Set RstRetenciones = CreditosDB.OpenRecordset("RetencionesTempo", dbOpenDynaset, dbSeeChanges)
  
  ' CargoCapital Prestado
        RstRetenciones.AddNew
        RstRetenciones!ConceptoCuotaId = 2 ' Capital Prestado
        RstRetenciones!TempoImporte = CapitalPrestado '(DSum("[Capital]", "TempoLiquidacion"))
        RstRetenciones!CodigoConcepto = 1
        RstRetenciones!SucursalID = Forms!Menu.SucursalID
        RstRetenciones!CreditoID = 0
        RstRetenciones!PcIP = Forms!Menu.IPPc
        RstRetenciones.Update
         
  
  ' Cargo Gastos Originacion
       If RstLinea!LineaFondocredito > 0 Then
          RstRetenciones.AddNew
          RstRetenciones!ConceptoCuotaId = 22 ' Gastos Originacion
          RstRetenciones!TempoImporte = Redondeo2(CapitalPrestado * 0.015)
          RstRetenciones!CodigoConcepto = 1
          Retenciones = Retenciones + RstRetenciones!TempoImporte
          RstRetenciones!SucursalID = Forms!Menu.SucursalID
          RstRetenciones!CreditoID = 0
          RstRetenciones!PcIP = Forms!Menu.IPPc
          RstRetenciones.Update
       End If
       
       If RstLinea!LineaFondocredito > 0 Then
          RstRetenciones.AddNew
          RstRetenciones!ConceptoCuotaId = 34 ' IVA Gastos Originacion
          RstRetenciones!TempoImporte = Redondeo2(Redondeo2(CapitalPrestado * 0.015) * 0.21)
          RstRetenciones!CodigoConcepto = 1
          Retenciones = Retenciones + RstRetenciones!TempoImporte
          RstRetenciones!SucursalID = Forms!Menu.SucursalID
          RstRetenciones!CreditoID = 0
          RstRetenciones!PcIP = Forms!Menu.IPPc
          RstRetenciones.Update
       End If
       
 
   '  Ajuste 1º cuota
       If difDias1eraCuota > 0 Then
             RstRetenciones.AddNew
             RstRetenciones!ConceptoCuotaId = 9 '2 'Ajuste 1º cuota
             RstRetenciones!TempoImporte = AjustePrimerCuota
             RstRetenciones!CodigoConcepto = 1
             RstRetenciones!Dias = Trim(difDias1eraCuota) + " días"
             RstRetenciones!SucursalID = Forms!Menu.SucursalID
             RstRetenciones!CreditoID = 0
             RstRetenciones!PcIP = Forms!Menu.IPPc
             
'             '****** Cargo ajuste 1era cuota al cuadro *****
'
'                    Tempo.AddNew
'                    Tempo!Ncuota = 1
'                    Tempo!fechavto = Forms![002AltaCreditos]!PrimerVencimiento ' Format(DateAdd("d", difDias1eraCuota, Now), "dd/mm/yyyy")
'                    Tempo!Tasa = Vartasa
'                    Tempo!TasaNominalAnual = Vartasa
'                    Tempo!Solicitado = CapitalPrestado
'                    Tempo!Prestado = CapitalPrestado
'                    Tempo!Capital = 0
'                    Tempo!Interes = AjustePrimerCuota
'                    Tempo!IvaInteres = AjustePrimerCuota * RstLinea!LineaIVASobreInt / 100
'                    Tempo!SdoCap = CapitalPrestado
'                    Tempo!SeguroVida = AjustePrimerCuota * 0.01
'                    Tempo!FondoGarantia = 0
'                    Tempo!Gastos = 0 '(Capital + Interes) * (RstLinea!LineaGastosAdmin / 100)
'                    Tempo!IVAGastos = Tempo!Gastos * IVAGeneral / 100
'                    Tempo!Cuota = AjustePrimerCuota + Tempo!IvaInteres + Tempo!SeguroVida ' Redondeo2(Capital + Tempo!Interes + Tempo!IvaInteres + Tempo!Gastos + Tempo!IVAGastos + Tempo!SeguroVida)
'                    Tempo!linea = Forms![002AltaCreditos]!ListaLineasHabilitadas.Column(1)
'                    If Not IsMissing(Forms![002AltaCreditos]![002subaltacreditossolicitante]!NombreII) Then
'                       Tempo!Cliente = Forms![002AltaCreditos]![002subaltacreditossolicitante]!NombreII
'                       Tempo!Documento = Forms![002AltaCreditos]![002subaltacreditossolicitante]!NumeroDoc
'                    End If
'                    Tempo!Sellado = 0 'SelladoII
'                    Tempo.Update
'
             
             
             '***********************************************
             Retenciones = Retenciones + RstRetenciones!TempoImporte
             RstRetenciones.Update
       End If
       
        '  IVA Ajuste 1º cuota
       If difDias1eraCuota > 0 Then
             RstRetenciones.AddNew
             RstRetenciones!ConceptoCuotaId = 36 ' IVA Ajuste 1º cuota
             RstRetenciones!TempoImporte = AjustePrimerCuota * (RstLinea!LineaIVAAjuste / 100)
             RstRetenciones!CodigoConcepto = 1
             RstRetenciones!Dias = Trim(difDias1eraCuota) + " días"
             RstRetenciones!SucursalID = Forms!Menu.SucursalID
             RstRetenciones!CreditoID = 0
             RstRetenciones!PcIP = Forms!Menu.IPPc
             Retenciones = Retenciones + RstRetenciones!TempoImporte
             RstRetenciones.Update
       End If
       
     ' Fondo Seguro
      If RstLinea!LineaFondocredito > 0 Then
          RstRetenciones.AddNew
          RstRetenciones!ConceptoCuotaId = 47 ' FONDO SEGURO
          RstRetenciones!TempoImporte = (DSum("[Capital]", "TempoLiquidacion") + DSum("[Interes]", "TempoLiquidacion")) * (RstLinea!LineaSeguroVida / 100)
          RstRetenciones!CodigoConcepto = 1
          Retenciones = Retenciones + RstRetenciones!TempoImporte
          RstRetenciones!SucursalID = Forms!Menu.SucursalID
          RstRetenciones!CreditoID = 0
          RstRetenciones!PcIP = Forms!Menu.IPPc
          RstRetenciones.Update
      End If
       
       ' Cargo sellado
       If RstLinea!LineaSellado > 0 Then
          RstRetenciones.AddNew
          RstRetenciones!ConceptoCuotaId = 24 ' Sellado
          RstRetenciones!TempoImporte = (DSum("[Capital]", "TempoLiquidacion") + DSum("[Interes]", "TempoLiquidacion")) * (RstLinea!LineaSellado / 100)
          RstRetenciones!CodigoConcepto = 1
          Retenciones = Retenciones + RstRetenciones!TempoImporte
          RstRetenciones!SucursalID = Forms!Menu.SucursalID
          RstRetenciones!CreditoID = 0
          RstRetenciones!PcIP = Forms!Menu.IPPc
          RstRetenciones.Update
       End If
       
       
       Forms![002AltaCreditos].Retenciones = Retenciones

'        ' Corrijo Seguro de Credito
'        DoCmd.RunSQL ("UPDATE TempoLiquidacion SET TempoLiquidacion.SeguroVida = 0, TempoLiquidacion.FondoGarantia = 0")
'        'Recalculo la cuota de todos los meses
'        Set Tempo = CurrentDb.OpenRecordset("TempoLiquidacion")
'        Tempo.MoveFirst
'        While Not Tempo.EOF
'              Tempo.Edit
'              If Forms![002Altacreditos].AniosFin < 70 Then
'                 Tempo!Cuota = Tempo!Capital + Tempo!Interes + Tempo!SeguroVida
'              Else
'                 Tempo!Cuota = Tempo!Capital + Tempo!Interes + Tempo!FondoGarantia
'              End If
'              Tempo.Update
'
'              Tempo.MoveNext
'        Wend
        
        Forms![002AltaCreditos].CapitalIII = CapitalPrestado '(DSum("[Capital]", "TempoLiquidacion"))
        'Forms![002Altacreditos].ACobrar = Forms![002Altacreditos].CapitalIII - Retenciones
        
        If Forms![002AltaCreditos].DiasHastaInicio < 0 Then
           Forms![002AltaCreditos].DiasHastaInicio = 0
        End If

        VarRetenciones = 0
        Forms![002AltaCreditos].ImporteCuota = DLookup("Cuota", "TempoLiquidacion")
        Forms![002AltaCreditos].MuestroRetenciones.SourceObject = "SubMuestroRetenciones"
        Dim DatosRetenciones As String
        DatosRetenciones = "SELECT RetencionesTempo.CreditoId, RetencionesTempo.CodigoConcepto, RetencionesTempo.ConceptoCuotaID, RetencionesTempo.CreditoId, RetencionesTempo.TempoImporte, RetencionesTempo.Dias, RetencionesTempo.PcIP, ConceptosCuota.ConceptoCuotaDescrip " & _
                           "FROM ConceptosCuota INNER JOIN RetencionesTempo ON ConceptosCuota.ConceptoCuotaID = RetencionesTempo.ConceptoCuotaID " & _
                           "WHERE (((RetencionesTempo.CodigoConcepto)=1) AND ((RetencionesTempo.ConceptoCuotaID)>2) AND ((RetencionesTempo.CreditoId)=0) AND ((RetencionesTempo.PcIP)='" & Trim(Forms!Menu!IPPc) & "'))"
        
        Forms![002AltaCreditos]![MuestroRetenciones].Form.RecordSource = DatosRetenciones
        If Forms![002AltaCreditos].AniosFin < 70 Then
           VarConceptoCuota = 4
        Else
           VarConceptoCuota = 5
        End If
        

                 
        DatoscargosORI = "SELECT TempConsCredAlta.CreditoID, TempConsCredAlta.SucursalCod, TempConsCredAlta.CreditoCuenta, TempConsCredAlta.CreditoFechaProxVto, TempConsCredAlta.CreditoSdoCap, TempConsCredAlta.EstadoCreditoID, TempConsCredAlta.OficinaCod, TempConsCredAlta.empleoid " & _
                         "FROM TempConsCredAlta " & _
                         "WHERE (((TempConsCredAlta.EstadoCreditoID)=4 Or (TempConsCredAlta.EstadoCreditoID)=8 Or (TempConsCredAlta.EstadoCreditoID)=10) AND ((TempConsCredAlta.OficinaCod)=107 Or (TempConsCredAlta.OficinaCod)=200 Or (TempConsCredAlta.OficinaCod)=111 Or (TempConsCredAlta.OficinaCod)=113 Or (TempConsCredAlta.OficinaCod)=100 Or (TempConsCredAlta.OficinaCod)=45) AND ((TempConsCredAlta.empleoid)=" & [Forms]![002AltaCreditos]![EmpleoSolicitanteID] & "))"
                 
        'LiquidaBoletaCob Forms!menu!CreditoId, Me.MarcoOpcionCobranza, Me.CuotasACobrar, Me.PorcBonificacion, Now
        Dim RstCargosORI As Recordset
        Set RstCargosORI = CreditosDB.OpenRecordset(DatoscargosORI, dbOpenDynaset, dbSeeChanges)
        DoCmd.RunSQL ("Delete * from TempGestionCobranza")
        If Not RstCargosORI.EOF Then
            RstCargosORI.MoveFirst
            'DoCmd.RunSQL ("Delete * from TempGestionCobranza")
            While Not RstCargosORI.EOF
                LiquidaBoletaCob RstCargosORI!CreditoID, 4, 100, 0, Now
                GraboComprobanteCargo
                RstCargosORI.MoveNext
            Wend
        End If
        'datoscargos = "SELECT TempConsCredAlta.CreditoID, TempConsCredAlta.SucursalCod, TempConsCredAlta.CreditoCuenta, TempConsCredAlta.CreditoFechaProxVto, TempConsCredAlta.CreditoSdoCap, TempConsCredAlta.EstadoCreditoID, TempConsCredAlta.OficinaCod, TempConsCredAlta.empleoid, Sum(TempGestionCobranza.ImporteConcepto) AS SumaDeImporteConcepto " & _
                      "FROM TempConsCredAlta INNER JOIN TempGestionCobranza ON TempConsCredAlta.CreditoID = TempGestionCobranza.CreditoId " & _
                      "GROUP BY TempConsCredAlta.CreditoID, TempConsCredAlta.SucursalCod, TempConsCredAlta.CreditoCuenta, TempConsCredAlta.CreditoFechaProxVto, TempConsCredAlta.CreditoSdoCap, TempConsCredAlta.EstadoCreditoID, TempConsCredAlta.OficinaCod, TempConsCredAlta.empleoid " & _
                      "HAVING (((TempConsCredAlta.EstadoCreditoID)=4 Or (TempConsCredAlta.EstadoCreditoID)=8 Or (TempConsCredAlta.EstadoCreditoID)=10) AND ((TempConsCredAlta.OficinaCod)=107 Or (TempConsCredAlta.OficinaCod)=200 Or (TempConsCredAlta.OficinaCod)=111 Or (TempConsCredAlta.OficinaCod)=113 Or (TempConsCredAlta.OficinaCod)=100 Or (TempConsCredAlta.OficinaCod)=45) AND ((TempConsCredAlta.empleoid)=" & [Forms]![002AltaCreditos]![EmpleoSolicitanteId] & "))"
                      
        datoscargos = "SELECT TempConsCredAlta.CreditoID, TempConsCredAlta.SucursalCod, TempConsCredAlta.CreditoCuenta, TempConsCredAlta.CreditoFechaProxVto, TempConsCredAlta.CreditoSdoCap, TempConsCredAlta.EstadoCreditoID, TempConsCredAlta.OficinaCod, TempConsCredAlta.empleoid, Sum(TempGestionCobranza.ImporteConcepto) AS SumaDeImporteConcepto, TempGestionCobranza.ComprobanteID " & _
                      "FROM TempConsCredAlta INNER JOIN TempGestionCobranza ON TempConsCredAlta.CreditoID = TempGestionCobranza.CreditoId " & _
                      "GROUP BY TempConsCredAlta.CreditoID, TempConsCredAlta.SucursalCod, TempConsCredAlta.CreditoCuenta, TempConsCredAlta.CreditoFechaProxVto, TempConsCredAlta.CreditoSdoCap, TempConsCredAlta.EstadoCreditoID, TempConsCredAlta.OficinaCod, TempConsCredAlta.empleoid, TempGestionCobranza.ComprobanteID " & _
                      "HAVING (((TempConsCredAlta.EstadoCreditoID)=4 Or (TempConsCredAlta.EstadoCreditoID)=8 Or (TempConsCredAlta.EstadoCreditoID)=10) AND ((TempConsCredAlta.OficinaCod)=107 Or (TempConsCredAlta.OficinaCod)=200 Or (TempConsCredAlta.OficinaCod)=111 Or (TempConsCredAlta.OficinaCod)=113 Or (TempConsCredAlta.OficinaCod)=100 Or (TempConsCredAlta.OficinaCod)=45) AND ((TempConsCredAlta.empleoid)=" & [Forms]![002AltaCreditos]![EmpleoSolicitanteID] & "))"

                 
        'DatosCargosSolicitante
        Dim RstCargos As Recordset
        Set RstCargos = CreditosDB.OpenRecordset(datoscargos, dbOpenDynaset, dbSeeChanges)
        '******************* Armo cargos completos *****************
        PasoATraves datoscargos, "TempoCargosCred"
        
        
        If RstCargos.RecordCount > 0 Then
            Forms![002AltaCreditos]![MuestroCargo].Form.RecordSource = datoscargos
            'Forms![002AltaCreditos]![MuestroCargo].Requery
            'MsgBox "aqui"
            Forms![002AltaCreditos]!Cargo = DSum("SumaDeImporteConcepto", "TempoCargosCred")
        Else
            'Forms![002Altacreditos]![MuestroCargo].Form.RecordSource = datoscargos
            Forms![002AltaCreditos].Cargo = 0
        End If
'End If
         


        'DoCmd.OpenQuery "CorrijoCapitalPrestado" ' Corrijo en TempoRetenciones (Remota)
'        DoCmd.RunSQL "UPDATE RetencionesTempo SET RetencionesTempo.TempoImporte = [Formularios]![002AltaCreditos]![CapitalIII] " + _
                     "WHERE RetencionesTempo.ConceptoCuotaID = 2 AND RetencionesTempo.PcIP = '" + Trim(Forms!Menu!IPPc) + _
                     "' AND RetencionesTempo.CreditoId = 0"


        'DoCmd.OpenQuery "CorrijoCapitalPrestadoII" 'Corrijo en TempoLiquidacion (Local)
        'DoCmd.RunSQL "UPDATE TempoLiquidacion SET TempoLiquidacion.Prestado = " + Trim(Forms![002Altacreditos]!CapitalIII)

        Forms![002AltaCreditos].MuestroRetenciones.Requery
        Forms![002AltaCreditos].MuestroCargo.Requery
        'A = [CapitalIII] - [Retenciones]

        'Busco Cargo
'        Dim Linea As Recordset
'        Dim Cargo As Currency
        
        'Forms![002Altacreditos].ImporteCuota = DLookup("Cuota", "TempoLiquidacion")
        'Forms![002Altacreditos].MuestroCargo.Requery
        
End If
Exit Sub

ErrFrances:
    MsgBox Err.description ' "Ocurrio un Error Al Generar Calculo Frances , Comunique CSI", vbCritical, "Sist. Gral de Creditos"


End Sub