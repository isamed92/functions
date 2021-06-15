Option Compare Database

Private Sub AbreConsulta_Click()
    'msgbox ""
    'DoCmd.OpenForm "001_consultaGralCreditos"
End Sub

Private Sub Calculo_Click()
On Error GoTo ErrCalculo

Dim FactorMes As Integer
Dim FechaFin As Date
Dim MargenSol As Currency
Dim MargenGar As Currency

        
        'MsgBox Me.ListaLineasHabilitadas.Column(8)
        If Me.ListaLineasHabilitadas.Column(3) = 109 Then
            IngNovedades.visible = True
        Else
            IngNovedades.visible = False
        End If
        Me.Cargo = 0
        
        Me.Guardar.Caption = "Grabar Credito"
'        If Me.Guardar.Caption = "Nuevo Turno" Then
'            DoCmd.Close
'            DoCmd.OpenForm "000AdminTurnos"
'            Exit Sub
'        End If
        
        If Me.ListaLineasHabilitadas.Column(8) = 1 And Me.OpGarante = 1 Then
            MsgBox "Esta Linea Requiere Garante", vbCritical, "Atención !!!!"
            Exit Sub
        
        
        End If
        
        MargenSol = IIf(IsNull(Me.MargenSolicitante), "0", Me.MargenSolicitante)
        MargenGar = IIf(IsNull(Me.MargenGarante), "0", Me.MargenGarante)
        If Me.OpGarante = 1 Then
            Me.MargenComputable = MargenSol
        Else
            Me.MargenComputable = IIf(MargenSol <= MargenGar, MargenSol, MargenGar)
        End If
        'Me.MargenComputable = IIf(MargenSol <= MargenGar, MargenSol, MargenGar)
        
        
        If DCount("EstadoCreditoID", "TempConsCredAlta", "EstadoCreditoID = 2") > 0 Then
           MsgBox "No se Puede Liquidar Crédito con otro crédito en Tramite Interno.", vbCritical, "Atención!!!"
           Exit Sub
        End If
        
        If DCount("EstadoCreditoID", "TempConsCredAlta", "EstadoCreditoID = 3") > 0 Then
           If MsgBox("El Cliente Ya Posee Crédito en Pago Programado." & NewLine & "Desea Liquidar Igualmente ?", vbQuestion + vbYesNo + vbDefaultButton2, "Atención!!!") = vbNo Then
                Exit Sub
           End If
        End If
        
        If DCount("EstadoCreditoID", "TempConsCredAlta", "EstadoCreditoID = 22") > 0 Then
           If MsgBox("El Cliente Ya Posee Crédito Activo Provisorio para Pago" & NewLine & "Desea Liquidar Igualmente ?", vbQuestion + vbYesNo + vbDefaultButton2, "Atención!!!") = vbNo Then
                Exit Sub
           End If
        End If
        
        If DCount("EstadoCreditoID", "TempConsCredAlta", "EstadoCreditoID = 23") > 0 Then
           If MsgBox("El Cliente Ya Posee Crédito Pendiente Autorización Superior" & NewLine & "Desea Liquidar Igualmente ?", vbQuestion + vbYesNo + vbDefaultButton2, "Atención!!!") = vbNo Then
                Exit Sub
           End If
        End If
        
        If IsNull(Me.ListaLineasHabilitadas) Or Me.ListaLineasHabilitadas = 0 Then
           MsgBox "Por favor, indique la línea de crédito.", vbCritical, "Atención!!!"
           Me.Capital.SetFocus
           Me.Plazo.SetFocus
           Exit Sub
        End If
        
        If Me.Plazo = 0 Or IsNull(Me.Plazo) Then
           MsgBox "Por favor, elija un plazo.", vbCritical, "Atención!!!"
           Me.Capital.SetFocus
           Me.Plazo.SetFocus
           Exit Sub
        
        End If
        
        If Me.Plazo > Val(Me.ListaLineasHabilitadas.Column(4)) Then
           MsgBox "El plazo indicado supera el plazo máximo de " + Trim(Me.ListaLineasHabilitadas.Column(4)) + " meses.", vbCritical, "Atención!!!"
           Me.Capital.SetFocus
           Me.Plazo.SetFocus
           Exit Sub
        End If
        
        If Me.PersonaId = 0 Or IsNull(Me.PersonaId) Then
           MsgBox "Por favor, indique el solicitante.", vbCritical, "Atención!!!"
           Me.Capital.SetFocus
           Me.Plazo = 0
           Exit Sub
        End If
        
        Me.Tasa = DameTasa(Me.ListaLineasHabilitadas, Me.Plazo)
        If Me.Tasa = 0 Then
           MsgBox "El plazo está fuera de los rangos para la línea. No puedo encontrar la tasa.", vbCritical, "Atención!!!"
           Me.Capital.SetFocus
           Me.Plazo = Null
           Me.Plazo.SetFocus
           Exit Sub
        End If
        
        Me.PlazoII = Me.Plazo
                
        '***** Comienzo Control Irregularidad Previa ******
        
        'Calculo automático la fecha del primer vencimiento
        Me.PrimerVencimiento = Det1erVTO(Me.ListaLineasHabilitadas, Now)
        FechaFin = str(DateAdd("m", Me.Plazo, Me.PrimerVencimiento))
        Me.FinCredito = FechaFin
        
                   
        'Determino Edad del Solicitante al terminar el credito
        AgregoBorroObservaciones 32, "Borro"
        If Not IsNull(Forms![002AltaCreditos]![002subaltacreditossolicitante]!Nacimiento) Then
           Me.AniosFin = DameEdad(Forms![002AltaCreditos]![002subaltacreditossolicitante]!Nacimiento, FechaFin)
           If Me.AniosFin < 18 Then
              AgregoBorroObservaciones 32
           End If
           If Me.AniosFin > 90 Then
              AgregoBorroObservaciones 32
           End If
        End If
        
        'Contolo Fecha Inicio Solicitante
        AgregoBorroObservaciones 6, "Borro"
        If IsNull(Forms![002AltaCreditos]![002subaltacreditossolicitante]!FechaInicio) Then
           AgregoBorroObservaciones 6
          
        End If
        
        
        AgregoBorroObservaciones 21, "Borro"
        If Me.Plazo > Val(Me.ListaLineasHabilitadas.Column(4)) Then
           AgregoBorroObservaciones 21
        End If
        
        ' Controlo Monto tope , grabo observacion
        AgregoBorroObservaciones 1, "Borro"
'        If Me.Capital > Val(Me.ListaLineasHabilitadas.Column(5)) Then
'           AgregoBorroObservaciones 1
'        End If
        
        
        If IsNull(Me.PrimerVencimiento) And (Me.Plazo = 0 Or IsNull(Me.Plazo)) Then
           MsgBox "Por favor, indique el primer vencimiento.", vbCritical, "Atención!!!"
           Me.PrimerVencimiento.SetFocus
           Exit Sub
        End If
        
        
        Dim DiasEnMes As Integer
        DiasEnMes = 30 ' Day(DateSerial(Year(Date), Month(Date) + 1, 0)) 'Tomo los días del mes siguiente
        Me.DiasHastaInicio = DateAdd("M", -1, Me.PrimerVencimiento) - Date
'        If Forms![002AltaCreditos]![002subaltacreditossolicitante]!ListaEmpleador = 292 Then
'            Me.DiasHastaInicio = 14
'        End If
        If Me.ListaLineasHabilitadas.Column(0) = 216 Then
            Me.DiasHastaInicio = 0
        End If
'*******************************Comienza Calculo ****************************************
'****************************************************************************************
 
 Select Case Me.ListaLineasHabilitadas.Column(7)
       Case 1
              FrancesTerceros Me.ListaLineasHabilitadas.Column(0), Me.Capital, Me.Plazo, Me.DiasHastaInicio
       Case 2
              FrancesANSES Me.ListaLineasHabilitadas.Column(0), Me.Capital, Me.Plazo, Me.DiasHastaInicio
       Case 3
              MsgBox "Sistema No Contemplado"
              Exit Sub
       Case 4
               Frances Me.ListaLineasHabilitadas.Column(0), Me.Capital, Me.Plazo, Me.DiasHastaInicio
       Case 5
               FrancesTerceros Me.ListaLineasHabilitadas.Column(0), Me.Capital, Me.Plazo, Me.DiasHastaInicio
       Case 6
               FrancesTerceros Me.ListaLineasHabilitadas.Column(0), Me.Capital, Me.Plazo, Me.DiasHastaInicio
       Case 7
               FondoCreditos Me.ListaLineasHabilitadas.Column(0), Me.Capital, Me.Plazo, Me.DiasHastaInicio
       Case 8
               'FondoCreditosLiquido Me.ListaLineasHabilitadas.Column(0), Me.Capital, Me.Plazo, Me.DiasHastaInicio
               FondoCreditos Me.ListaLineasHabilitadas.Column(0), Me.Capital, Me.Plazo, Me.DiasHastaInicio
       Case 9
               'FondoCreditosCF Me.ListaLineasHabilitadas.Column(0), Me.Capital, Me.Plazo, Me.DiasHastaInicio
               FondoCreditosCFCargo Me.ListaLineasHabilitadas.Column(0), Me.Capital, Me.Plazo, Me.DiasHastaInicio
       Case 10
               'FondoCreditosLiquido Me.ListaLineasHabilitadas.Column(0), Me.Capital, Me.Plazo, Me.DiasHastaInicio
               FondoCreditosCFL Me.ListaLineasHabilitadas.Column(0), Me.Capital, Me.Plazo, Me.DiasHastaInicio
       Case 11
               FondoCreditos201 Me.ListaLineasHabilitadas.Column(0), Me.Capital, Me.Plazo, Me.DiasHastaInicio
       Case 12
               of202_ResXXX Me.ListaLineasHabilitadas.Column(0), Me.Capital, Me.Plazo, Me.DiasHastaInicio
       Case 13
               of202_ResXXXX Me.ListaLineasHabilitadas.Column(0), Me.Capital, Me.Plazo, Me.DiasHastaInicio
       Case 14
               of203_Res1085 Me.ListaLineasHabilitadas.Column(0), Me.Capital, Me.Plazo, Me.DiasHastaInicio
       Case 15
               of202_ROMANO Me.ListaLineasHabilitadas.Column(0), Me.Capital, Me.Plazo, Me.DiasHastaInicio
 End Select
'*****************************************************************************************
'*****************************************************************************************
       
       '******* Inicio controles Post Liquidacion *****************
        
        Me.MuestroRetenciones.Requery
        Me.MuestroCargo.Requery
        'Me.Cargo = Forms![002AltaCreditos]![MuestroCargo]!totalcargo
        Me.ACobrar = Me.CapitalIII - Me.Retenciones - Me.Cargo
        
        'Me Fijo si le queda Algo a Cobrar
        If Me.ACobrar < 0 Then
           MsgBox "El monto solicitado es menor al cargo ", vbCritical, "Atención!!!"
        End If
                
        'Me Tope Linea
        AgregoBorroObservaciones 42, "Borro"
        If Me.ACobrar > (Val(Me.ListaLineasHabilitadas.Column(5)) + 1) Then
            AgregoBorroObservaciones 42
            Me.Guardar.enabled = True
        Else
            Me.Guardar.enabled = True
        End If
        
                
        'Me fijo en el márgen solicitante
        AgregoBorroObservaciones 2, "Borro"
        If Me.MargenComputable < Me.ImporteCuota Then
           AgregoBorroObservaciones 2, "Le falta $" & str(Redondeo(Me.ImporteCuota - Me.MargenComputable))
        End If
        
        'Me fijo en el márgen garante
        AgregoBorroObservaciones 7, "Borro"
        If Me.OpGarante = 2 Then
           If Me.MargenGarante < Me.ImporteCuota Then
              AgregoBorroObservaciones 7, "Le falta $" & str(Redondeo(Me.ImporteCuota - Me.MargenGarante))
           End If
        End If
        
        'Me fijo si tiene fecha Cese en el empleo
        AgregoBorroObservaciones 17, "Borro"
        If Not IsNull(Forms![002AltaCreditos]![002subaltacreditossolicitante]!FechaFin) Then
           AgregoBorroObservaciones 17, "Borro"
           If Forms![002AltaCreditos]![002subaltacreditossolicitante]!FechaFin < Me.FinCredito Then
              AgregoBorroObservaciones 17
           End If
        End If
        
        If Forms![002AltaCreditos]![002subaltacreditossolicitante]!RelacionLaboral = 1 Then
           AgregoBorroObservaciones 17, "Borro"
    
        End If
                        
        'Me fijo si el trabajo elegido no es planta permanente y además no tiene garante
        AgregoBorroObservaciones 4, "Borro"
        If (Forms![002AltaCreditos]![002subaltacreditossolicitante]!RelacionLaboral > 1 And Forms![002AltaCreditos]![002subaltacreditossolicitante]!RelacionLaboral <> 6) And (Me.GaranteID = 0 Or IsNull(Me.GaranteID)) Then
           AgregoBorroObservaciones 4
        End If
        Me.Irregularidades.Requery
        MsgBox "Credito Liquidado !!!"
        
                
        'Me.Guardar.enabled = True
        
'*******************************  fin Calculo    ****************************************
        
        Exit Sub
ErrCalculo:
        MsgBox Err.description, vbCritical, "Atención!!!"
End Sub

Private Sub Capital_AfterUpdate()
    
    If Me.Capital < 0 Then
       MsgBox "El Capital no debe ser menor a 0 !!!", vbCritical, "Atención !!!"
       Me.Capital = 0
       Exit Sub
    End If
    If Me.Plazo > 0 Then
        MsgBox "Atención , Va a Reliquidar el Crédito ", vbInformation, "  ...::: Sistema Gral. de Créditos :::..."
        'Calculo_Click
        Plazo_AfterUpdate
    End If
    Me.Plazo.SetFocus

'    If Me.Capital > Val(Me.ListaLineasHabilitadas.Column(5)) Then
'        MsgBox "El Monto Maximo para esta linea es : " & Val(Me.ListaLineasHabilitadas.Column(5)), vbCritical, "Atención !!!"
'        Me.Capital = Val(Me.ListaLineasHabilitadas.Column(5))
'        Exit Sub
'    End If
    
       
End Sub

Private Sub CargaDatos_Click()
On Error GoTo ErrCargaDatos
Dim RstCreditos As Recordset
Dim Cadena As String
Set RstCreditos = CreditosDB.OpenRecordset("SELECT Creditos.* FROM Creditos " & _
                                           "WHERE (((Creditos.CreditoID)=" & Me.AuxCreditoID & "))")

    Me.EmpleoSolicitanteID = RstCreditos!EmpleoId
    Me.EmpleoGaranteID = RstCreditos!CreditoEmpleoIdGarante
    Me.GaranteID = RstCreditos!CreditoGaranteId
    Me.MargenComputable = RstCreditos!CreditoMargenComputable
    
    ' Cargo Solicitante
    
    Forms![002AltaCreditos]![002subaltacreditossolicitante]!EmpleoId = Me.EmpleoSolicitanteID
    Forms![002AltaCreditos]![002subaltacreditossolicitante]!EmpleosSolicitante.RowSource = ""
    Cadena = "SELECT Empleadores.EmpleadorDescrip, Empleos.* " & _
             "FROM Empleadores INNER JOIN Empleos ON Empleadores.EmpleadorID = Empleos.EmpleadorID " & _
             "WHERE (((Empleos.PersonaID)=" & Me.PersonaId & "))"
            
    PasoATraves Cadena, "AAATempoEmpleoSolicitante"
    Forms![002AltaCreditos]![002subaltacreditossolicitante]!EmpleosSolicitante.RowSource = "SELECT AAATempoEmpleoSolicitante.EmpleadorDescrip as 'Lugar de Trabajo', AAATempoEmpleoSolicitante.EmpleoLegajo as 'Legajo', AAATempoEmpleoSolicitante.Empleoid FROM AAATempoEmpleoSolicitante"
    Forms![002AltaCreditos]![002subaltacreditossolicitante]!EmpleosSolicitante.Requery
    
    DoCmd.RunSQL "DELETE * FROM TempoDeducciones"
    DoCmd.RunSQL "INSERT INTO TempoDeducciones ( DeduccionId, DeduccionImporte, TipoCliente, DeduccionNombre ) " & _
                 "SELECT EmpleoMargenDeducciones.DeduccionId, EmpleoMargenDeducciones.EmpMargDeducImporte, 1 AS Expr1, DeduccionesMargen.DeduccionNombre " & _
                 "FROM EmpleoMargenDeducciones INNER JOIN DeduccionesMargen ON EmpleoMargenDeducciones.DeduccionId = DeduccionesMargen.DeduccionId " & _
                 "WHERE (((EmpleoMargenDeducciones.EmpleoMargenID)=" & RstCreditos!CreditoEmpleoMargenID & "))"
    
    Forms![002AltaCreditos]![002subaltacreditossolicitante]!Elegidos.Requery
    Dim RstMargen As Recordset
    Set RstMargen = CreditosDB.OpenRecordset("SELECT EmpleoMargen.* FROM EmpleoMargen " & _
                                              "WHERE (((EmpleoMargen.EmpleoMargenId)=" & RstCreditos!CreditoEmpleoMargenID & "))", dbOpenDynaset, dbSeeChanges)
    
    Forms![002AltaCreditos]![002subaltacreditossolicitante]!MargenID = RstCreditos!CreditoEmpleoMargenID
    Forms![002AltaCreditos]![002subaltacreditossolicitante]!IngresoMensual = RstMargen!EmpleoMargenNeto
    Forms![002AltaCreditos]![002subaltacreditossolicitante]!CertTrabajo = RstMargen!EmpleoMargenMargen
    Forms![002AltaCreditos]![002subaltacreditossolicitante]!CuotasActuales = RstMargen!EmpleoMargenImpCuotaSolic
    Forms![002AltaCreditos]![002subaltacreditossolicitante]!ComoGaranteSolicitante = RstMargen!EmpleoMargenImpCuotaGarante
    Forms![002AltaCreditos]![002subaltacreditossolicitante]!MargenComputable.Requery
    
    'requisitos
    DoCmd.RunSQL "DELETE * FROM TempoRequisitos"
    DoCmd.RunSQL "INSERT INTO TempoRequisitos ( RequisitoID, Tipo ) " & _
                 "SELECT CreditoRequisitos.RequisitoID, CreditoRequisitos.RequisitoCreditoQuien " & _
                 "FROM CreditoRequisitos " & _
                 "WHERE (((CreditoRequisitos.CreditoID)=" & Me.AuxCreditoID & ") AND ((CreditoRequisitos.RequisitoCreditoQuien)=1))"

    Forms![002AltaCreditos]![002subaltacreditossolicitante]!Presentados.Requery
    
    ' Garante
    If Not Me.GaranteID = 0 Then
        Me.OpGarante = 2
        OpGarante_AfterUpdate
        Dim cade As String
        ' Cargo datos Garante
        cade = "SELECT Personas.* FROM Personas WHERE (((Personas.PersonaID)=" & Me.GaranteID & "))"
           PasoATraves cade, "AAATempoGarante"
            
           Set RstDatPersonales = CreditosDB.OpenRecordset("AAATempoGarante")
           If Not RstDatPersonales.EOF Then
                Forms![002AltaCreditos]![002subaltacreditosgarante]!PersonaId = Me.GaranteID
                Forms![002AltaCreditos]![002subaltacreditosgarante]!TipoDoc = RstDatPersonales!TipoDocumentoID
                Forms![002AltaCreditos]![002subaltacreditosgarante]!NumeroDoc = RstDatPersonales!PersonaNumDoc
                Forms![002AltaCreditos]![002subaltacreditosgarante]!CUIT = RstDatPersonales!PersonaCUIT
                Forms![002AltaCreditos]![002subaltacreditosgarante]!NombreII = RstDatPersonales!PersonaDescrip
                Forms![002AltaCreditos]![002subaltacreditosgarante]!TipoPersona = RstDatPersonales!PersonaTipoPersona
                Forms![002AltaCreditos]![002subaltacreditosgarante]!Sexo = RstDatPersonales!PersonaSexo
                Forms![002AltaCreditos]![002subaltacreditosgarante]!Nacimiento = RstDatPersonales!PersonaFechaNac
                Forms![002AltaCreditos]![002subaltacreditosgarante]!Nacionalidad = RstDatPersonales!NacionalidadID
                Forms![002AltaCreditos]![002subaltacreditosgarante]!EstadoCivil = RstDatPersonales!EstadoCivilID
                Forms![002AltaCreditos]![002subaltacreditosgarante]!Telefono = RstDatPersonales!PersonaTelContacto
                Forms![002AltaCreditos]![002subaltacreditosgarante]!TelefonoII = RstDatPersonales!PersonaTelMovil
                Forms![002AltaCreditos]![002subaltacreditosgarante]!Email = RstDatPersonales!PersonaEmail
                Forms![002AltaCreditos]![002subaltacreditosgarante]!SituacionImpositiva = RstDatPersonales!SituacionImpositivaID
                Forms![002AltaCreditos]![002subaltacreditosgarante]!EstadoPersona = RstDatPersonales!EstadoPersonaId
                'domi = "SELECT TiposDomicilios.TipoDomicilioDescrip, Domicilios.* " & _
                       "FROM Domicilios INNER JOIN TiposDomicilios ON Domicilios.TipoDomicilioID = TiposDomicilios.TipoDomicilioID " & _
                       "WHERE (((Domicilios.PersonaId) " & Me.GaranteId & "])) " & _
                       "ORDER BY Domicilios.TipoDomicilioID"
                'Forms![002Altacreditos]![002subaltacreditosgarante]!ListaDomSolicitante.RowSource = "SELECT TiposDomicilios.TipoDomicilioDescrip, Domicilios.* " & _
                       "FROM Domicilios INNER JOIN TiposDomicilios ON Domicilios.TipoDomicilioID = TiposDomicilios.TipoDomicilioID " & _
                       "WHERE (((Domicilios.PersonaId) " & Me.GaranteId & "])) " & _
                       "ORDER BY Domicilios.TipoDomicilioID"
                Forms![002AltaCreditos]![002subaltacreditosgarante]!ListaDomSolicitante.Requery
                
        End If
        
        
                
            Forms![002AltaCreditos]![002subaltacreditosgarante]!EmpleoId = Me.EmpleoGaranteID
            Forms![002AltaCreditos]![002subaltacreditosgarante]!EmpleosSolicitante.RowSource = ""
            Cadena = "SELECT Empleadores.EmpleadorDescrip, Empleos.* " & _
                     "FROM Empleadores INNER JOIN Empleos ON Empleadores.EmpleadorID = Empleos.EmpleadorID " & _
                     "WHERE (((Empleos.PersonaID)=" & Me.GaranteID & "))"
                    
            PasoATraves Cadena, "AAATempoEmpleosGarante"
            Forms![002AltaCreditos]![002subaltacreditosgarante]!EmpleosSolicitante.RowSource = "SELECT AAATempoEmpleosGarante.EmpleadorDescrip as 'Lugar de Trabajo', AAATempoEmpleosGarante.EmpleoLegajo as 'Legajo', AAATempoEmpleosGarante.Empleoid FROM AAATempoEmpleosGarante"
            Forms![002AltaCreditos]![002subaltacreditosgarante]!EmpleosSolicitante.Requery
            
            'margen
            
            'DoCmd.RunSQL "DELETE * FROM TempoDeducciones"
            DoCmd.RunSQL "INSERT INTO TempoDeducciones ( DeduccionId, DeduccionImporte, TipoCliente, DeduccionNombre ) " & _
                         "SELECT EmpleoMargenDeducciones.DeduccionId, EmpleoMargenDeducciones.EmpMargDeducImporte, 2 AS Expr1, DeduccionesMargen.DeduccionNombre " & _
                         "FROM EmpleoMargenDeducciones INNER JOIN DeduccionesMargen ON EmpleoMargenDeducciones.DeduccionId = DeduccionesMargen.DeduccionId " & _
                         "WHERE (((EmpleoMargenDeducciones.EmpleoMargenID)=" & RstCreditos!CreditoEmpleoMargenIDGarante & "))"
            
            Forms![002AltaCreditos]![002subaltacreditosgarante]!Elegidos.Requery
            'Dim RstMargenG As Recordset
            Set RstMargen = CreditosDB.OpenRecordset("SELECT EmpleoMargen.* FROM EmpleoMargen " & _
                                                      "WHERE (((EmpleoMargen.EmpleoMargenId)=" & RstCreditos!CreditoEmpleoMargenIDGarante & "))", dbOpenDynaset, dbSeeChanges)
            
            Forms![002AltaCreditos]![002subaltacreditosgarante]!MargenID = RstCreditos!CreditoEmpleoMargenID
            Forms![002AltaCreditos]![002subaltacreditosgarante]!IngresoMensual = RstMargen!EmpleoMargenNeto
            Forms![002AltaCreditos]![002subaltacreditosgarante]!CertTrabajo = RstMargen!EmpleoMargenMargen
            Forms![002AltaCreditos]![002subaltacreditosgarante]!CuotasActuales = RstMargen!EmpleoMargenImpCuotaSolic
            Forms![002AltaCreditos]![002subaltacreditosgarante]!ComoGaranteSolicitante = RstMargen!EmpleoMargenImpCuotaGarante
            Forms![002AltaCreditos]![002subaltacreditosgarante]!MargenComputable.Requery
            
            'requisitos
            DoCmd.RunSQL "DELETE * FROM TempoRequisitos"
            DoCmd.RunSQL "INSERT INTO TempoRequisitos ( RequisitoID, Tipo ) " & _
                         "SELECT CreditoRequisitos.RequisitoID, CreditoRequisitos.RequisitoCreditoQuien " & _
                         "FROM CreditoRequisitos " & _
                         "WHERE (((CreditoRequisitos.CreditoID)=" & Me.AuxCreditoID & ") AND ((CreditoRequisitos.RequisitoCreditoQuien)=2))"
        
            Forms![002AltaCreditos]![002subaltacreditosgarante]!Presentados.Requery
            
        
            
                
        
        
    Else
        Me.OpGarante = 1
        OpGarante_AfterUpdate
    End If
    
RstCreditos.Close

    
Exit Sub
ErrCargaDatos:
    MsgBox Err.description
End Sub


Private Sub CesionEmp_Click()
    DoCmd.OpenReport "CesionHaberesCPA", acViewPreview, , , acDialog
    
    DoCmd.OpenReport "AutorizacionPago", acViewPreview, , , acDialog
End Sub

Private Sub Cuadro_Click()

        Dim Tempo As Recordset
        Set Tempo = CreditosDB.OpenRecordset("TempoLiquidacion")
        If Not Tempo.EOF Then
           DoCmd.OpenForm "MuestroCuadro", , , , , , "TempoLiquidacion"
        Else
           MsgBox "El cuadro de marcha no tiene información.", vbCritical
        End If
        Set Tempo = Nothing


End Sub

Private Sub Form_Current()
    SetFormIcon hWnd, "\\134.14.14.9\neocred\Produccion\Imagenes\logocred.ico"
End Sub

Private Sub Form_Load()
   
    Select Case ControlPermiso(Forms!Menu.GrupoId, "2")
               Case 0
                    MsgBox "Lo siento, usted no tiene permiso para ejecutar esta función.", vbCritical, "Atención!!!"
                    DoCmd.Close
                    Exit Sub
               'Case 1
               '     Me.Grabar.enabled = False
               'Case 2
               '     Me.Grabar.enabled = True
        End Select
    
    
    Me.[002_SubAltaCreditos].SetFocus
    Me.OpGarante = 1
    Me.TabGarante.visible = False
    
        
    
End Sub

Private Sub Guardar_Click()
On Error GoTo ErrGuardarCred
        Dim Tempo As Recordset
        'Inicio controles
        
        If Me.Guardar.Caption = "Nuevo Turno" Then
            DoCmd.Close
            DoCmd.OpenForm "000AdminCallCenter"
            Exit Sub
        End If
        
        If Me.Guardar.Caption = "Actualizar Corrección" Then
            MsgBox "No Tiene Permisos Suficientes", vbCritical
            'DoCmd.Close
            'DoCmd.OpenForm "000AdminTurnos"
            Exit Sub
        End If
                
        If Me.PersonaId = 0 Or IsNull(Me.PersonaId) Then
           MsgBox "Elija un cliente.", vbCritical, "Error!!!"
           'Me.Documento.SetFocus
           Exit Sub
        End If
        
        'Controlo que el solicitante tengo un domicilio "Real"
        If Not DameDomicilioReal(Me.PersonaId) Then
           MsgBox "El cliente no tiene un domicilio 'Real'. Por favor, controle.", vbCritical, "Atención!!!"
        Exit Sub
            
        End If
        
        If Me.AniosHoy < 18 Then
           MsgBox "El cliente no tiene la edad suficiente para operar. Por favor, controle.", vbCritical, "Atención!!!"
           Exit Sub
        End If
        
        'If Me.CreditoId > 0 Then
        '   MsgBox "Ya se grabó un crédito, por favor cancele para poder grabar uno nuevo.", vbCritical, "Atención!!!"
        '   Exit Sub
        'End If
        
        
        If Me.ACobrar < 0 Then
           MsgBox "El neto a cobrar no puede ser negativo.", vbCritical, "Atención!!!"
           Exit Sub
               
        End If
        
        If IsNull(Me.PersonaId) Or Me.PersonaId = 0 Then
           MsgBox "Por favor, indique el cliente.", vbCritical, "Atención!!!"
          'Me.Documento.SetFocus
           Exit Sub
        End If
        
        If Me.PersonaId = Me.GaranteID Then
           MsgBox "El solicitante no puede ser el garante.", vbCritical, "Atención!!!"
           'Me.Documento.SetFocus
           Exit Sub
        End If
        
        'Controlo el empleo del solicitante
'        Set TempoEmpleos = CreditosDB.OpenRecordset("Select * From TempoEmpleos Where TipoCliente = 1 and TempoEleccion = True")
'        If Not TempoEmpleos.EOF Then
'           TempoEmpleos.MoveLast
'           If TempoEmpleos.RecordCount > 1 Then
'              MsgBox "Usted no puede elegir más de un empleo para cesionar.", vbCritical, "Atención!!!"
'              Exit Sub
'           End If
'        Else ' No eligió ningún empleo
'           MsgBox "Por favor, elija un empleo para cesionar.", vbCritical, "Atención!!!"
'           Exit Sub
'
'        End If
        
'        If Me.OpGarante = 2 Then 'Tiene garante, me fijo si elegió un empleo de este
'
'           Set TempoEmpleos = CreditosDB.OpenRecordset("Select * From TempoEmpleos Where TipoCliente = 2 and TempoEleccion = True")
'           If Not TempoEmpleos.EOF Then
'              TempoEmpleos.MoveLast
'              If TempoEmpleos.RecordCount > 1 Then
'                 MsgBox "Usted no puede elegir más de un empleo del garante.", vbCritical, "Atención!!!"
'                 Exit Sub
'              End If
'           Else ' No eligió ningún empleo
'              MsgBox "Por favor, elija un empleo del garante.", vbCritical, "Atención!!!"
'              Exit Sub
'           End If
'        End If
'        Set TempoEmpleos = Nothing
        
        If Me.ImporteCuota = 0 Then
           MsgBox "Por favor, indique un crédito a otorgar.", vbCritical, "Atención!!!"
           Exit Sub
        End If
        
        'Mefijo que tenga grabado un domicilio
        Dim domi As Recordset
        Set domi = CurrentDb.OpenRecordset("Select PersonaId From Domicilios Where PersonaId = " + Trim(Me.PersonaId))
        If domi.EOF Then
           MsgBox "El cliente no tiene grabado un domicilio.", vbCritical
           Exit Sub
        End If
        
        'Fin controles
        
        
        If MsgBox("Está usted seguro de grabar el crédito?", 36, "Atención!!!") = vbNo Then
           Exit Sub
        End If
        
        
        Dim VarCuenta As Long
        Dim VarGastos, VarSellado, VarInteres, VarCargoAnses As Currency
        Dim VarCesion As Integer
        VarCesion = 256
        VarGastos = 0
        VarSellado = 0
        VarCargoAnses = 0
        
        VarInteres = DSum("[Interes]", "TempoLiquidacion")
        
        VarSellado = DSum("[TempoImporte]", "RetencionesTempo", "PcIP = '" + Trim(Forms!Menu.IPPc) + "' AND CreditoId = 0 AND ConceptoCuotaID = 24")
       
       Dim Filtro As QueryDef
       Dim Resultado As Recordset
       Dim ParametrosAltaCreditos As String
       Set Filtro = CurrentDb.CreateQueryDef("")
       Filtro.Connect = CurrentDb.TableDefs("Creditos").Connect
              
Dim ParaAltaCred1 As String
Dim ParaAltaCred2 As String
ParaAltaCred1 = Trim(Me.PersonaId) & ", " & _
                Trim(Me.ListaLineasHabilitadas) & ", " & _
                Trim(Forms!Menu.SucursalID) & ", " & _
                Trim(Me.ListaLineasHabilitadas.Column(3)) & ", " & _
                str(VarCuenta) & ", " & _
                str(Me.Capital) & ", " & _
                Trim(Me.Plazo) & ", " & _
                str(Me.CapitalIII) & ", " & _
                "'" & FechaInsertTo(Date) & "'" & ", " & _
                "'" & FechaInsertTo(Me.PrimerVencimiento) & "'" & ", " & _
                str(VarGastos) & ", " & _
                IIf(IsNull(str(VarSellado)), "0", str(VarSellado)) & ", " & _
                str(Me.ImporteCuota) & ", " & _
                str(Me.CapitalIII) & ", " & _
                str(VarInteres) & ", " & _
                Trim(Me.Plazo) & ", " & _
                Trim(Forms![002AltaCreditos]![002subaltacreditossolicitante]!ListaEmpleador) & ", " & _
                str(VarCesion) & ", " & _
                str(Me.Tasa) & ", " & _
                Trim(Forms!Menu!UsuarioID) & ", " & _
                str(Me.ACobrar) & ", "
ParaAltaCred2 = Trim(IIf(IsNull(Me.GaranteID), "0", Me.GaranteID)) & ", " & _
                str(IIf(IsNull(Me.CFTEA), "0", Me.CFTEA)) & ", " & _
                Trim(Me.EmpleoSolicitanteID) & ", " & _
                Trim(IIf(IsNull(Me.EmpleoGaranteID), "0", Me.EmpleoGaranteID)) & ", " & _
                str(IIf(IsNull(Me.Cargo), "0", Me.Cargo)) & ", " & _
                "'" & Trim(Forms!Menu.IPPc) & "'" & ", " & _
                "'" & Trim(Forms!Menu.NombreDelPc) & "'" & ", " & _
                str(IIf(IsNull(Forms![002AltaCreditos]![002subaltacreditossolicitante]!MargenComputable), "0", Forms![002AltaCreditos]![002subaltacreditossolicitante]!MargenComputable)) & ", " & _
                str(IIf(IsNull(Me.Factor), "0", Me.Factor)) & ", " & _
                str(VarCargoAnses) & ", " & _
                str(IIf(IsNull(Forms![002AltaCreditos]![002subaltacreditossolicitante]!IngresoMensual), "0", Forms![002AltaCreditos]![002subaltacreditossolicitante]!IngresoMensual)) & ", " & _
                str(IIf(IsNull(Forms![002AltaCreditos]![002subaltacreditossolicitante]!Neto), "0", Forms![002AltaCreditos]![002subaltacreditossolicitante]!Neto)) & ", " & _
                str(IIf(IsNull(Forms![002AltaCreditos]![002subaltacreditossolicitante]!MargenComp), "0", Forms![002AltaCreditos]![002subaltacreditossolicitante]!MargenComp)) & ", " & _
                str(IIf(IsNull(Forms![002AltaCreditos]![002subaltacreditossolicitante]!MargenComputable), "0", Forms![002AltaCreditos]![002subaltacreditossolicitante]!MargenComputable)) & ", "
                   
                If Me.OpGarante = 1 Then
                     ParaAltaCred3 = "0, 0, 0, 0,"
                Else
                ParaAltaCred3 = str(IIf(Forms![002AltaCreditos]![002subaltacreditosgarante]!IngresoMensual = "", "0", Forms![002AltaCreditos]![002subaltacreditosgarante]!IngresoMensual)) & ", " & _
                                str(IIf(IsNull(Forms![002AltaCreditos]![002subaltacreditosgarante]!Neto), "0", Forms![002AltaCreditos]![002subaltacreditosgarante]!Neto)) & ", " & _
                                str(IIf(IsNull(Forms![002AltaCreditos]![002subaltacreditosgarante]!MargenComp), "0", Forms![002AltaCreditos]![002subaltacreditosgarante]!MargenComp)) & ", " & _
                                str(IIf(IsNull(Forms![002AltaCreditos]![002subaltacreditosgarante]!MargenComputable), "0", Forms![002AltaCreditos]![002subaltacreditosgarante]!MargenComputable)) & ", "
                End If
       ParaAltaCred4 = str(IIf(IsNull(Forms![002AltaCreditos]![002subaltacreditossolicitante]!MargenID), "0", Forms![002AltaCreditos]![002subaltacreditossolicitante]!MargenID)) & ", " & _
                       str(IIf(IsNull(Forms![002AltaCreditos]![002subaltacreditosgarante]!MargenID), "0", Forms![002AltaCreditos]![002subaltacreditosgarante]!MargenID))
                
       Filtro.SQL = "execute AgregoCredito " & ParaAltaCred1 & ParaAltaCred2 & ParaAltaCred3 & ParaAltaCred4
       Set Resultado = Filtro.OpenRecordset()

       'Me.CreditoId = Resultado.Fields(0)
        
            
       'Me.CreditoId = Creditos!CreditoId
       'Me.Cuenta = VarCuenta
        
        
       'Forms!Menu.AuxCreditoId = Me.CreditoId
       
       
       Forms!Menu.AuxCreditoID = Resultado.Fields(0)
        
        
        DoCmd.OpenQuery "CompletoTempoRetenciones"
        Dim TempoRet As Recordset
        Set TempoRet = CurrentDb.OpenRecordset("Select * From RetencionesTempo Where PcIP = '" + Trim(Forms!Menu.IPPc) + _
                                                                               "' AND CreditoId = 0", dbOpenDynaset, dbSeeChanges)
        If Not TempoRet.EOF Then
           TempoRet.MoveFirst
           While Not TempoRet.EOF
                 TempoRet.Edit
                 TempoRet!CreditoID = Forms!Menu.AuxCreditoID
                 TempoRet.Update
                 TempoRet.MoveNext
           Wend
        End If
        TempoRet.Close
        
        
        Dim TipoResponsabilidad As Integer
        TipoResponsabilidad = 1
        
        'Grabo en tabla IntegrantesCreditos
        DoCmd.RunSQL "INSERT INTO IntegrantesCreditos (CreditoId, PersonaId, TipoResponsabilidadID, EmpleoID )" + _
                     "Select " + Trim(Forms!Menu.AuxCreditoID) + ", " + Trim(Me.PersonaId) + ", " + _
                     str(TipoResponsabilidad) + ", " + Trim(Me.EmpleoSolicitanteID)
                     

        If Me.OpGarante = 2 Then 'Tiene garante
           TipoResponsabilidad = 2
           DoCmd.RunSQL "INSERT INTO IntegrantesCreditos (CreditoId, PersonaId, TipoResponsabilidadID, EmpleoID )" + _
                        "Select " + Trim(Forms!Menu.AuxCreditoID) + ", " + Trim(Me.GaranteID) + ", " + _
                        str(TipoResponsabilidad) + ", " + Trim(Me.EmpleoGaranteID)
           
        End If

        DoCmd.OpenQuery "LlenoCuadrosMarcha"
            
        'Grabo CreditoObservaciones
        Set Tempo = CreditosDB.OpenRecordset("TempoIrregularidades")
        If Not Tempo.EOF Then
           Dim CreditoObs As Recordset
           Set CreditoObs = CreditosDB.OpenRecordset("CreditoObservaciones", dbOpenDynaset, dbSeeChanges)
           Tempo.MoveFirst
           While Not Tempo.EOF
                 CreditoObs.AddNew
                 CreditoObs!IrregularidadId = Tempo!IrregularidadId
                 CreditoObs!CreditoID = Forms!Menu.AuxCreditoID
                 CreditoObs!CreditoObsObservaciones = Tempo!Observaciones
                 CreditoObs.Update
                 Tempo.MoveNext
           Wend
           CreditoObs.Close
        End If
            
            
            'Grabo en CreditoRequisitos
            Set Tempo = CreditosDB.OpenRecordset("TempoRequisitos")
            If Not Tempo.EOF Then
               Dim CreditosReq As Recordset
               Set CreditosReq = CreditosDB.OpenRecordset("CreditoRequisitos", dbOpenDynaset, dbSeeChanges)
               Tempo.MoveFirst
               While Not Tempo.EOF
                     CreditosReq.AddNew
                     CreditosReq!CreditoID = Forms!Menu.AuxCreditoID
                     CreditosReq!RequisitoId = Tempo!RequisitoId
                     CreditosReq!RequisitoCreditoQuien = Tempo!tipo
                     CreditosReq.Update
                     Tempo.MoveNext
               Wend
            End If
            Set Tempo = CreditosDB.OpenRecordset("TempoRequisitosGarante")
            If Not Tempo.EOF Then
               'Dim CreditosReq As Recordset
               Set CreditosReq = CreditosDB.OpenRecordset("CreditoRequisitos", dbOpenDynaset, dbSeeChanges)
               Tempo.MoveFirst
               While Not Tempo.EOF
                     CreditosReq.AddNew
                     CreditosReq!CreditoID = Forms!Menu.AuxCreditoID
                     CreditosReq!RequisitoId = Tempo!RequisitoId
                     CreditosReq!RequisitoCreditoQuien = Tempo!tipo
                     CreditosReq.Update
                     Tempo.MoveNext
               Wend
            End If
            
            
            'Grabo Cargos. Desnormalizo, grabo el id del crédito nuevo, los ids de los que se hicieron cargo y el total
            'DoCmd.OpenQuery "CompletoTempoCreditosPersona"
            Dim ImporteCargos As Currency
            Dim TiraCargos As String
            Dim SrtTempo As String
            
            'SrtTempo = "SELECT TempConsCredAlta.CreditoID, TempConsCredAlta.SucursalCod, TempConsCredAlta.CreditoCuenta, TempConsCredAlta.CreditoFechaProxVto, TempConsCredAlta.CreditoSdoCap, TempConsCredAlta.EstadoCreditoID, TempConsCredAlta.OficinaCod, TempConsCredAlta.empleoid " & _
                       "FROM TempConsCredAlta " & _
                       "WHERE (((TempConsCredAlta.EstadoCreditoID)=4) AND ((TempConsCredAlta.OficinaCod)=107 Or (TempConsCredAlta.OficinaCod)=113 Or (TempConsCredAlta.OficinaCod)=100) AND ((TempConsCredAlta.empleoid)=" & [Forms]![002AltaCreditos]![EmpleoSolicitanteID] & "))"
            
            SrtTempo = "SELECT TempoCargosCred.CreditoID, TempoCargosCred.SucursalCod, TempoCargosCred.CreditoCuenta, TempoCargosCred.CreditoFechaProxVto, TempoCargosCred.CreditoSdoCap, TempoCargosCred.EstadoCreditoID, TempoCargosCred.OficinaCod, TempoCargosCred.empleoid, TempoCargosCred.SumaDeImporteConcepto, TempoCargosCred.ComprobanteID " & _
                       "FROM TempoCargosCred " & _
                       "WHERE (((TempoCargosCred.empleoid)=" & [Forms]![002AltaCreditos]![EmpleoSolicitanteID] & "))"

            Set Tempo = CreditosDB.OpenRecordset(SrtTempo)


            If Not Tempo.EOF Then
               Dim CargoCredito As Recordset
               Set CargoCredito = CreditosDB.OpenRecordset("CreditoCargos", dbOpenDynaset, dbSeeChanges)
               Tempo.MoveFirst
               While Not Tempo.EOF
                     CargoCredito.AddNew
                     CargoCredito!CreditoID = Forms!Menu.AuxCreditoID
                     CargoCredito!CreditoCargoCreditoId = Tempo!CreditoID
                     CargoCredito!CreditoCargoImporte = Tempo!SumaDeImporteConcepto
                     CargoCredito!CreditoCargoEstado = "P"
                     CargoCredito!ComprobanteID = Tempo!ComprobanteID
                     CargoCredito.Update
                     ImporteCargos = ImporteCargos + Tempo!SumaDeImporteConcepto
                     TiraCargos = TiraCargos + "*" + DameCasaOfCta(Tempo!CreditoID)
                     Tempo.MoveNext
               Wend
               CargoCredito.Close
            End If
            
            
            'Grabo CreditoRetenciones
            Dim CreditosRet As Recordset
            Set Tempo = CreditosDB.OpenRecordset("SELECT RetencionesTempo.*, RetencionesTempo.PcIP AS Crit " & _
                                                 "FROM RetencionesTempo " & _
                                                 "WHERE (((RetencionesTempo.PcIP)='" & Trim(Forms!Menu.IPPc) & "'))", dbOpenDynaset, dbSeeChanges)
            If Not Tempo.EOF Then
               Set CreditosRet = CreditosDB.OpenRecordset("CreditoRetenciones", dbOpenDynaset, dbSeeChanges)
               Tempo.MoveFirst
               While Not Tempo.EOF
                     CreditosRet.AddNew
                     CreditosRet!CreditoID = Forms!Menu.AuxCreditoID
                     CreditosRet!ConceptoCuotaID = Tempo!ConceptoCuotaID
                     If Tempo!ConceptoCuotaID = 2 Then
                        CreditosRet!CredRetSigno = 1
                     End If
                     CreditosRet!CredRetImporte = Tempo!TempoImporte
                     CreditosRet!CredRetDias = Tempo!Dias
                     CreditosRet.Update
                     Tempo.MoveNext
               Wend
               
'               If ImporteCargos > 0 Then 'Agrego registro de Cargo en Credito
'                  CreditosRet.AddNew
'                  CreditosRet!CreditoId = Forms!menu.AuxCreditoID
'                  CreditosRet!ConceptoCuotaID = 26
'                  CreditosRet!CredRetImporte = ImporteCargos
'                  CreditosRet!CredRetDias = TiraCargos
'                  CreditosRet.Update
'               End If
'
'               If ImporteCargos > 0 Then 'Agrego registro de capital prestado
'                  CreditosRet.AddNew
'                  CreditosRet!CreditoId = Forms!menu.AuxCreditoID
'                  CreditosRet!ConceptoCuotaID = 2
'                  CreditosRet!CredRetImporte = Me.CapitalIII
'                  CreditosRet!CredRetDias = ""
'                  CreditosRet.Update
'
'               End If
               CreditosRet.Close
            End If
            Tempo.Close
            
           
        'Grabo en CreditoCuotasPendientes
        DoCmd.OpenQuery "LlenoCreditoCuotasPendientes"
        
        'Grabo en CreditoDeducciones
        DoCmd.OpenQuery "LlenoCreditoDeducciones"
        DoCmd.OpenQuery "BorroTempoCreditosPersona"
        DoCmd.OpenQuery "LlenoTempoCreditosPersona"
        'Me.Antecedentes.SourceObject = "AntecedentesCrediticios"
        'CommitTrans
        
        'Grabo es CuadroCreditos el cuadro de marcha
        ImpactoCuadro Forms!Menu.AuxCreditoID
        
        Me.Guardar.Caption = "Nuevo Turno"
                
        '********* Imprimo Documentacion  **********
        If MsgBox("Crédito Grabado, listo para Autorización !!!" & vbNewLine & "Desea imprimir documentación ???", vbYesNo + vbInformation, "Felicitaciones !!!") = vbYes Then
            ImprimeSartaOtorgamientos Forms!Menu.AuxCreditoID, Me.OpGarante, Me.ListaLineasHabilitadas
        End If
        
        If MsgBox("Desea imprimir contrato VISA ???", vbYesNo + vbInformation, "CONTRATO VISA") = vbYes Then
           DoCmd.OpenReport "ContratoVisa", acViewNormal
           DoCmd.OpenReport "ContratoVisa1", acViewNormal
           DoCmd.OpenReport "ContratoVisa2", acViewNormal
           DoCmd.OpenReport "ContratoVisa3", acViewNormal
           DoCmd.OpenReport "ContratoVisa4", acViewNormal
           DoCmd.OpenReport "ContratoVisa5", acViewNormal
           DoCmd.OpenReport "ContratoVisa6", acViewNormal
           DoCmd.OpenReport "ContratoVisa7", acViewNormal
           DoCmd.OpenReport "ContratoVisa8", acViewNormal
           DoCmd.OpenReport "ContratoVisa9", acViewNormal
           DoCmd.OpenReport "ContratoVisa10", acViewNormal
           DoCmd.OpenReport "ContratoVisa11", acViewNormal
           DoCmd.OpenReport "ContratoVisa12", acViewNormal
                      
        End If
        
        Exit Sub
ErrGuardarCred:
        MsgBox "Ocurrió un error, no se grabó el crédito.", vbCritical, "Atención!!!"
        MsgBox Err.description & " Err.Number:" &  Err.Number &
        'Rollback
        
End Sub

Private Sub ImprimeLiquidacion_Click()
On Error GoTo errImpOtorg
        If Me.OpGarante = 2 Then
           DoCmd.OpenReport "SolicitudConGarante", acViewPreview
           DoCmd.OpenReport "ConsentimientoConGarante", acViewPreview
           
        Else
           DoCmd.OpenReport "Solicitud", acViewPreview
           DoCmd.OpenReport "ConsentimientoSolicitante", acViewPreview
        End If
        DoCmd.OpenReport "Consentimiento", acViewPreview
        DoCmd.OpenReport "ReporteMargen", acViewPreview
        
Exit Sub
errImpOtorg:
    MsgBox Err.description
End Sub

Private Sub IngNovedades_Click()
    DoCmd.OpenReport "IngresosNovANSES", acViewPreview, , , acDialog
End Sub

Private Sub OpGarante_AfterUpdate()

        
'        Me.Solapa.Pages(5).visible = False
'        Me.Solapa.Pages(6).visible = False
'        Me.Solapa.Pages(7).visible = False
'        Me.Solapa.Pages(8).visible = False
'        AgregoBorroObservaciones 20, "Borro"
'
'        If Me.Garante = 1 Then 'Con Garante
'           Me.Solapa.Pages(5).visible = True
'           Me.Solapa.Pages(6).visible = True
'           Me.Solapa.Pages(7).visible = True
'           Me.Solapa.Pages(8).visible = True
'           AgregoBorroObservaciones 20, "Le faltan todos"
'        Else
'           PrimeraVezGar = True
'           Me.NumeroDoc = 0
'           Me.GaranteId = 0
'           NumeroDoc_AfterUpdate
'           AgregoBorroObservaciones 32, "Borro"
'           If (Me.AuxRelacionLabolarId > 1 And Me.AuxRelacionLabolarId <> 6) Then
'              AgregoBorroObservaciones 32
'           End If
'        End If
On Error GoTo ErrGarante
    If Me.OpGarante = 1 Then
        Me.TabGarante.visible = False
    Else
        Me.TabGarante.visible = True
    End If
Exit Sub
ErrGarante:
    MsgBox Err.description
        
        
End Sub

Private Sub Plazo_AfterUpdate()
    
    If Me.Plazo < 0 Then
       Me.Plazo = 0
       Exit Sub
    End If

'    If Me.Plazo > Val(Me.ListaLineasHabilitadas.Column(4)) Then
'        MsgBox "El Tope de la linea es : " & Val(Me.ListaLineasHabilitadas.Column(4)), vbCritical
'        Me.Plazo = Val(Me.ListaLineasHabilitadas.Column(4))
'        Exit Sub
'    End If
   Calculo_Click
End Sub

