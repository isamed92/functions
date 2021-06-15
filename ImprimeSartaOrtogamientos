Public Sub ImprimeSartaOtorgamientos(CreditoID As Long, OpGarante As Integer, LineaCreditoID As Integer)

        Dim Reportes As Recordset
        Set Reportes = CurrentDb.OpenRecordset("OtorgamientoReportes", dbOpenDynaset, dbSeeChanges)
                
                'Cuadro de Marcha
                
                   DoCmd.OpenReport "CuadroMarchaLiquidacion", acViewNormal
                
                ' Solicitud
                If OpGarante = 2 Then
                   Reportes.AddNew
                   Reportes!CreditoID = Forms!Menu.AuxCreditoID
                   Reportes!OtoReporteCamino = GuardoPDF(Trim(Forms!Menu.AuxCreditoID), "SolicitudConGarante")
                   DoCmd.OpenReport "SolicitudConGarante"
                   Reportes.Update
                Else
                   Reportes.AddNew
                   Reportes!CreditoID = Forms!Menu.AuxCreditoID
                   Reportes!OtoReporteCamino = GuardoPDF(Trim(Forms!Menu.AuxCreditoID), "Solicitud")
                   DoCmd.OpenReport "Solicitud"
                   Reportes.Update
                   
                End If
                               
                ' Concentimiento
                If LineaCreditoID = 215 Then
                    Reportes.AddNew
                    Reportes!CreditoID = Forms!Menu.AuxCreditoID
                    Reportes!OtoReporteCamino = GuardoPDF(Trim(Forms!Menu.AuxCreditoID), "ConsentimientoHipodromoCasino")
                    DoCmd.OpenReport "ConsentimientoHipodromoCasino", , , , , Forms!Menu.AuxCreditoID
                    Reportes.Update
                Else
                    If LineaCreditoID = 216 Then
                        Reportes.AddNew
                        Reportes!CreditoID = Forms!Menu.AuxCreditoID
                        Reportes!OtoReporteCamino = GuardoPDF(Trim(Forms!Menu.AuxCreditoID), "l216Consentimiento")
                        DoCmd.OpenReport "l216Consentimiento", , , , , Forms!Menu.AuxCreditoID
                        Reportes.Update
                    Else
                        If LineaCreditoID = 217 Then
                        Reportes.AddNew
                        Reportes!CreditoID = Forms!Menu.AuxCreditoID
                        Reportes!OtoReporteCamino = GuardoPDF(Trim(Forms!Menu.AuxCreditoID), "l217Consentimiento")
                        DoCmd.OpenReport "l217Consentimiento", , , , , Forms!Menu.AuxCreditoID
                        Reportes.Update
                        Else
                            Reportes.AddNew
                            Reportes!CreditoID = Forms!Menu.AuxCreditoID
                            Reportes!OtoReporteCamino = GuardoPDF(Trim(Forms!Menu.AuxCreditoID), "Consentimiento")
                            DoCmd.OpenReport "Consentimiento", , , , , Forms!Menu.AuxCreditoID
                            Reportes.Update
                        End If
                    End If
                End If
                
                'Reporte Margen
                Reportes.AddNew
                Reportes!CreditoID = Forms!Menu.AuxCreditoID
                Reportes!OtoReporteCamino = GuardoPDF(Trim(Forms!Menu.AuxCreditoID), "ReporteMargen")
                DoCmd.OpenReport "ReporteMargen"
                Reportes.Update
                
                'Boleta liquidacion
                Reportes.AddNew
                Reportes!CreditoID = Forms!Menu.AuxCreditoID
                Reportes!OtoReporteCamino = GuardoPDF(Trim(Forms!Menu.AuxCreditoID), "BoletaLiquidacion")
                DoCmd.OpenReport "BoletaLiquidacion"
                Reportes.Update
                
                'Cesion de Haberes
                If LineaCreditoID = 216 Then
                    Reportes.AddNew
                    Reportes!CreditoID = Forms!Menu.AuxCreditoID
                    Reportes!OtoReporteCamino = GuardoPDF(Trim(Forms!Menu.AuxCreditoID), "L216CesionHaberes")
                    DoCmd.OpenReport "L216CesionHaberes"
                    Reportes.Update
                Else
                    If LineaCreditoID = 215 Then
                        Reportes.AddNew
                        Reportes!CreditoID = Forms!Menu.AuxCreditoID
                        Reportes!OtoReporteCamino = GuardoPDF(Trim(Forms!Menu.AuxCreditoID), "L215CesionHaberes")
                        DoCmd.OpenReport "L215CesionHaberes"
                        Reportes.Update
                    Else
                        If LineaCreditoID = 217 Then
                        Reportes.AddNew
                        Reportes!CreditoID = Forms!Menu.AuxCreditoID
                        Reportes!OtoReporteCamino = GuardoPDF(Trim(Forms!Menu.AuxCreditoID), "L217CesionHaberes")
                        DoCmd.OpenReport "L217CesionHaberes"
                        Reportes.Update
                        Else
                            Reportes.AddNew
                            Reportes!CreditoID = Forms!Menu.AuxCreditoID
                            Reportes!OtoReporteCamino = GuardoPDF(Trim(Forms!Menu.AuxCreditoID), "CesionHaberes")
                            DoCmd.OpenReport "CesionHaberes"
                            Reportes.Update
                        End If
                    End If
                End If
                
                'Declaracion Jurada
                Reportes.AddNew
                Reportes!CreditoID = Forms!Menu.AuxCreditoID
                Reportes!OtoReporteCamino = GuardoPDF(Trim(Forms!Menu.AuxCreditoID), "DeclaracionJurada")
                DoCmd.OpenReport "DeclaracionJurada"
                Reportes.Update
                
                'Pagares
                If OpGarante = 2 Then
                   If LineaCreditoID = 216 Then
                        Reportes.AddNew
                        Reportes!CreditoID = Forms!Menu.AuxCreditoID
                        Reportes!OtoReporteCamino = GuardoPDF(Trim(Forms!Menu.AuxCreditoID), "l216PagareGarante")
                        DoCmd.OpenReport "l216PagareGarante"
                        Reportes.Update
                   Else
                        If LineaCreditoID = 217 Then
                             Reportes.AddNew
                             Reportes!CreditoID = Forms!Menu.AuxCreditoID
                             Reportes!OtoReporteCamino = GuardoPDF(Trim(Forms!Menu.AuxCreditoID), "l217PagareGarante")
                             DoCmd.OpenReport "l217PagareGarante"
                             Reportes.Update
                        Else
                             Reportes.AddNew
                             Reportes!CreditoID = Forms!Menu.AuxCreditoID
                             Reportes!OtoReporteCamino = GuardoPDF(Trim(Forms!Menu.AuxCreditoID), "PagareGarante")
                             DoCmd.OpenReport "PagareGarante"
                             Reportes.Update
                        End If
                   End If
                   
                   If LineaCreditoID = 215 Then
                        Reportes.AddNew
                        Reportes!CreditoID = Forms!Menu.AuxCreditoID
                        Reportes!OtoReporteCamino = GuardoPDF(Trim(Forms!Menu.AuxCreditoID), "ContratoHipodromoCasino")
                        DoCmd.OpenReport "ContratoHipodromoCasino", , , , , Forms!Menu.AuxCreditoID
                        Reportes.Update
                   Else
                        If LineaCreditoID = 216 Then
                            Reportes.AddNew
                            Reportes!CreditoID = Forms!Menu.AuxCreditoID
                            Reportes!OtoReporteCamino = GuardoPDF(Trim(Forms!Menu.AuxCreditoID), "L216ContratoMutuo")
                            DoCmd.OpenReport "L216ContratoMutuo", , , , , Forms!Menu.AuxCreditoID
                            Reportes.Update
                        Else
                            Reportes.AddNew
                            Reportes!CreditoID = Forms!Menu.AuxCreditoID
                            Reportes!OtoReporteCamino = GuardoPDF(Trim(Forms!Menu.AuxCreditoID), "Contrato")
                            DoCmd.OpenReport "Contrato", , , , , Forms!Menu.AuxCreditoID
                            Reportes.Update
                        End If
                   End If
                   
                   Reportes.AddNew
                   Reportes!CreditoID = Forms!Menu.AuxCreditoID
                   Reportes!OtoReporteCamino = GuardoPDF(Trim(Forms!Menu.AuxCreditoID), "L216CesionHaberesGarante")
                   DoCmd.OpenReport "L216CesionHaberesGarante"
                   Reportes.Update
                   
                Else 'Pagare sin Garante
                                   
                   If LineaCreditoID = 216 Then
                        Reportes.AddNew
                        Reportes!CreditoID = Forms!Menu.AuxCreditoID
                        Reportes!OtoReporteCamino = GuardoPDF(Trim(Forms!Menu.AuxCreditoID), "l216PagareSolicitante")
                        DoCmd.OpenReport "l216PagareSolicitante"
                        Reportes.Update
                   Else
                        Reportes.AddNew
                        Reportes!CreditoID = Forms!Menu.AuxCreditoID
                        Reportes!OtoReporteCamino = GuardoPDF(Trim(Forms!Menu.AuxCreditoID), "PagareSolicitante")
                        DoCmd.OpenReport "PagareSolicitante"
                        Reportes.Update
                   End If
                   If LineaCreditoID = 215 Then
                        Reportes.AddNew
                        Reportes!CreditoID = Forms!Menu.AuxCreditoID
                        Reportes!OtoReporteCamino = GuardoPDF(Trim(Forms!Menu.AuxCreditoID), "ContratoHipodromoCasino")
                        DoCmd.OpenReport "ContratoHipodromoCasino", , , , , Forms!Menu.AuxCreditoID
                        Reportes.Update
                   Else
                        If LineaCreditoID = 216 Then
                            Reportes.AddNew
                            Reportes!CreditoID = Forms!Menu.AuxCreditoID
                            Reportes!OtoReporteCamino = GuardoPDF(Trim(Forms!Menu.AuxCreditoID), "L216ContratoMutuo")
                            DoCmd.OpenReport "L216ContratoMutuo", , , , , Forms!Menu.AuxCreditoID
                            Reportes.Update
                        Else
                            If LineaCreditoID = 217 Then
                                Reportes.AddNew
                                Reportes!CreditoID = Forms!Menu.AuxCreditoID
                                Reportes!OtoReporteCamino = GuardoPDF(Trim(Forms!Menu.AuxCreditoID), "L217ContratoMutuo")
                                DoCmd.OpenReport "L217ContratoMutuo", , , , , Forms!Menu.AuxCreditoID
                                Reportes.Update
                            Else
                                Reportes.AddNew
                                Reportes!CreditoID = Forms!Menu.AuxCreditoID
                                Reportes!OtoReporteCamino = GuardoPDF(Trim(Forms!Menu.AuxCreditoID), "Contrato")
                                DoCmd.OpenReport "Contrato", , , , , Forms!Menu.AuxCreditoID
                                Reportes.Update
                            End If
                        End If
                   End If
                                      
                End If

End Sub
