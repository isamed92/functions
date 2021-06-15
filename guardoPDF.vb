Public Function GuardoPDF(Prefijo As String, ReporteNombre As String) As String

       'Mando el reporte al PDF y lo guardo en el servidor
       DoCmd.OutputTo acReport, ReporteNombre, acFormatPDF, "\\134.14.14.9\NeoCred\Produccion\PDFs\" & Prefijo + ReporteNombre + ".PDF"
       'DoCmd.OutputTo acReport, ReporteNombre, acFormatPDF, "\\134.14.14.9\creditospdfs\" & Prefijo + ReporteNombre + ".PDF"
       
       'Devuelvo el camino para guardarlo en una tabla o para poder abrir el PDF inmediatamente
       'GuardoPDF = "\\134.14.14.9\creditospdfs\" & Prefijo + ReporteNombre + ".PDF"
       GuardoPDF = "\\134.14.14.9\NeoCred\Produccion\PDFs\" & Prefijo + ReporteNombre + ".PDF"
       
End Function