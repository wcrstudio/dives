Imports System.Windows.Forms
Public Class diversity
    Public LibCalculate As New decalculate.calculate
    Public Sub New()
        Main()
    End Sub

    Public Function ShannonWiener(ByRef oDGV As DataGridView) As String
        Dim lcol As Long, lrow As Long
        Dim iR As Long, iC As Integer
        Dim lCT As Long, lCel As Long
        Dim iCelValue As Integer
        Dim lFi As Double, lLogFi As Double
        Dim sLev As String, somaT As Double
        Dim sRTF As String = ""


        lcol = oDGV.ColumnCount
        lrow = oDGV.RowCount
        Try
            iLog = LibLoadConfig.LoadConfig("CONFIG", "LOGBASE", "2")
            If Auditoria(oDGV) = False Then
                If MsgAuditoria() = False Then
                    Return sRetAuditoria

                End If
            End If
            For iC = 0 To lcol - 1 Step 1

                lCT = 0
                For iR = 0 To lrow - 1 Step 1
                    If oDGV.Rows(iR).Cells(iC).Value <> String.Empty Then
                        lCT += CDbl(oDGV.Rows(iR).Cells(iC).Value)
                    End If

                Next iR
                sLev = oDGV.Columns(iC).HeaderText.ToString
                sRTF &= "<div class='row'><div class='col-xs-4 col-lg-4'>Levantamento: " & sLev & "</div>"
                sRTF &= "<div class='col-xs-4 col-lg-4'>Valor N: " & lCT & "</div>"
                sRTF &= "<div class='col-xs-4 col-lg-4'>Base Logarítmica: " & iLog & "</div></div>"
                sRTF &= "<div class='table-responsive'><table class='table table-bordered table-hover' ><thead><tr class='info'>" &
                    "<th>Espécie</th>" &
                    "<th>Número de Indivíduos</th>" &
                    "<th>Valor Fi</th>" &
                    "<th>Valor Log(Fi)</th>" &
                    "<th>Valor Fi x Log(Fi)</th>" &
                    "</tr></thead><tbody>"


                For iCelValue = 0 To lrow - 1
                    If oDGV.Rows(iCelValue).Cells(iC).Value <> String.Empty Then
                        lCel = CDbl(oDGV.Rows(iCelValue).Cells(iC).Value)
                    Else
                        lCel = 0
                    End If
                    sRTF &= "<tr>"
                    sRTF &= "<td>" & oDGV.Rows(iCelValue).HeaderCell.Value & "</td>"
                    sRTF &= "<td class='text-center'>" & lCel & "</td>"

                    lFi = lCel / lCT
                    If lFi > 0 Then ' 0 Then
                        sRTF &= "<td>" & Math.Round(lFi, 4) & "</td>"
                        lLogFi = Math.Log(lFi, iLog)

                        somaT += (lFi * lLogFi)
                        sRTF &= "<td class='text-center'>" & Format(lLogFi, "###,###,###,##0.0###") & "</td>"
                        sRTF &= "<td class='text-center'>" & Format((lFi * lLogFi), "###,###,###,##0.0###") & "</td>"
                    Else
                        sRTF &= "<td class='text-center'>0</td>"
                        sRTF &= "<td class='text-center'>0</td>"
                    End If
                    sRTF &= "</tr>"
                Next iCelValue
                sRTF &= "</tbody><tfooter><tr class='text-right text-info'>"
                sRTF &= "<td colspan=5>Índice Diversidade de Shannon-Wiener: " & Format(Math.Abs(somaT), "###,###,###,##0.0###") & "</td></tr></tfooter>"
                sRTF &= "</table></div><hr/>"
                somaT = 0
            Next iC

            sResult = sHTMLStart & "<h3 class='alert alert-success'>Relatório da Diversidade de Shannon-Wiener</h3>" & sRTF & sFooter & sHTMLEnd
            Return sResult

        Catch ex As Exception

            LibError.ErrorView(ex)
            Return retError(sHTMLStart, ex, sHTMLEnd)
        End Try


    End Function

    Public Function fDivSW(ByRef oDGV As DataGridView, ByVal iC As Integer) As Double

        Dim lcol As Long, lrow As Long
        Dim iR As Long
        Dim lCT As Long, lCel As Long
        Dim iCelValue As Integer
        Dim lFi As Double, lLogFi As Double
        Dim somaT As Double

        lcol = oDGV.ColumnCount
        lrow = oDGV.RowCount
        Try
            iLog = LibLoadConfig.LoadConfig("CONFIG", "LOGBASE", "2")
            If Auditoria(oDGV) = False Then
                If MsgAuditoria() = False Then
                    Return sRetAuditoria

                End If
            End If

            lCT = 0
            For iR = 0 To lrow - 1 ' Step 1
                If oDGV.Rows(iR).Cells(iC).Value <> String.Empty Then _
                lCT += oDGV.Rows(iR).Cells(iC).Value
            Next iR

            For iCelValue = 0 To lrow - 1
                If oDGV.Rows(iCelValue).Cells(iC).Value <> String.Empty Then
                    lCel = oDGV.Rows(iCelValue).Cells(iC).Value
                Else
                    lCel = 0
                End If

                If lCel > 0 Then
                    lFi = lCel / lCT
                    If lFi > 0 Then ' 0 Then

                        lLogFi = Math.Log(lFi, iLog)

                        somaT += (lFi * lLogFi)
                    End If
                End If

            Next iCelValue
            Return Math.Abs(somaT)
        Catch ex As Exception
            LibError.ErrorView(ex)
            Return 0
        End Try
    End Function

    Public Function divSimpson(ByRef oDGV As DataGridView) As String
        Dim lcol As Long, lrow As Long
        Dim iR As Long, iC As Integer
        Dim lCT As Long, lCel As Long
        Dim iCelValue As Integer
        Dim lFi As Double, lDivSim As Double
        Dim sLev As String, somaT As Double
        Dim lNvsN As Double, lDivSimTotal As Double
        Dim sRTF As String = ""
        lcol = oDGV.ColumnCount
        lrow = oDGV.RowCount
        Try
            If Auditoria(oDGV) = False Then
                If MsgAuditoria() = False Then
                    Return sRetAuditoria

                End If
            End If

            For iC = 0 To lcol - 1 Step 1
                lCT = 0
                For iR = 0 To lrow - 1 Step 1
                    If oDGV.Rows(iR).Cells(iC).Value <> String.Empty Then _
                    lCT += CDbl(oDGV.Rows(iR).Cells(iC).Value)
                Next iR

                sLev = oDGV.Columns(iC).HeaderText.ToString
                sRTF &= "<div class='row'><div class='col-xs-6 col-lg-6'>Levantamento: " & sLev & "</div>"
                sRTF &= "<div class='col-xs-6 col-lg-6'>Valor N: " & lCT & "</div></div>"
                sRTF &= "<div class='table-responsive'><table class='table table-bordered table-hover' ><thead><tr class='info'><th>Espécie</th><th>Número de Indivíduos</th><th>n*(n-1)</th></tr></thead><tbody>"
                lNvsN = lCT * (lCT - 1)
                For iCelValue = 0 To lrow - 1
                    If oDGV.Rows(iCelValue).Cells(iC).Value <> String.Empty Then
                        lCel = CDbl(oDGV.Rows(iCelValue).Cells(iC).Value)
                        sRTF &= "<tr>"
                        sRTF &= "<td>" & oDGV.Rows(iCelValue).HeaderCell.Value & "</td>"
                        sRTF &= "<td class='text-center'>" & lCel & "</td>"

                        lFi += lCel * (lCel - 1)
                        somaT += lFi
                        If lFi > 0 Then ' 0 Then
                            sRTF &= "<td class='text-center'>" & Format(lFi, "###,###,###,##0.0###") & "</td>"
                        Else
                            sRTF &= "<td>n*(n-1): 0</td>"
                        End If
                        sRTF &= "</tr>"
                    End If
                  

                Next iCelValue
                lDivSim = 1 - (lFi / lNvsN)
                lDivSimTotal += (lFi / lNvsN)
                
                sRTF &= "</tbody><tfooter><tr class='text-info'>"
                sRTF &= "<td colspan=3>Índice Diversidade de Simpson: " & Format(Math.Abs(lDivSim), "###,###,###,##0.0###") & "</td>"
                sRTF &= "</tr></tfooter>"
                sRTF &= "</table></div><hr/>"
                somaT = 0
                lFi = 0
            Next iC
            'sRTF &= "Índice Diversidade de Simpson Total: " & Format(Math.Abs(lDivSimTotal), "###,###,###,##0.0###") & vbNewLin

            sResult = sHTMLStart & "<h3 class='alert alert-success'>Relatório da Diversidade de Simpson</h3>" & sRTF & sFooter & sHTMLEnd
            Return sResult

        Catch ex As Exception

            LibError.ErrorView(ex)
            Return retError(sHTMLStart, ex, sHTMLEnd)
        End Try

    End Function

    Public Function fdivSimpson(ByRef oDGV As DataGridView, ByVal ic As Integer)
        Dim lcol As Long, lrow As Long
        Dim iR As Long
        Dim lCT As Long, lCel As Long
        Dim iCelValue As Integer
        Dim lFi As Double, lDivSim As Double
        Dim sLev As String, somaT As Double
        Dim lNvsN As Double
        lcol = oDGV.ColumnCount
        lrow = oDGV.RowCount
        Try
            If Auditoria(oDGV) = False Then
                If MsgAuditoria() = False Then
                    Return sRetAuditoria

                End If
            End If
            lCT = 0
            For iR = 0 To lrow - 1 Step 1
                If oDGV.Rows(iR).Cells(ic).Value <> String.Empty Then _
                lCT += CDbl(oDGV.Rows(iR).Cells(ic).Value)
            Next iR

            sLev = oDGV.Columns(ic).HeaderText.ToString

            lNvsN = lCT * (lCT - 1)
            For iCelValue = 0 To lrow - 1
                If oDGV.Rows(iCelValue).Cells(ic).Value <> String.Empty Then
                    lCel = CDbl(oDGV.Rows(iCelValue).Cells(ic).Value)
                Else
                    lCel = 0
                End If

               
                If lCel > 0 Then
                    lFi += lCel * (lCel - 1)
                    somaT += lFi
                End If
               

            Next iCelValue
            lDivSim = 1 - (lFi / lNvsN)
            Return lDivSim
        Catch ex As Exception
            LibError.ErrorView(ex)
            Return 0
        End Try
    End Function

    Public Function divMargalef(ByRef oDGV As DataGridView) As String
        Dim lcol As Long, lrow As Long
        Dim iR As Long, iC As Integer
        Dim lCT As Long, lCel As Long
        Dim iCelValue As Integer
        Dim lDivMargalef As Double, sLev As String
        Dim sp As Double, lDivMargalefTotal As Double
        Dim sRTF As String = ""
        lcol = oDGV.ColumnCount
        lrow = oDGV.RowCount
        Try
            If Auditoria(oDGV) = False Then
                If MsgAuditoria() = False Then
                    Return sRetAuditoria

                End If
            End If
            For iC = 0 To lcol - 1 Step 1
                lCT = 0
                For iR = 0 To lrow - 1 Step 1
                    If oDGV.Rows(iR).Cells(iC).Value <> String.Empty Then _
                    lCT += CDbl(oDGV.Rows(iR).Cells(iC).Value)
                Next iR

                sLev = oDGV.Columns(iC).HeaderText.ToString
                sRTF &= "<div class='row'><div class='col-xs-4 col-lg-4'>Levantamento: " & sLev & "</div>"
                sRTF &= "<div class='col-xs-4 col-lg-4'>Valor N: " & lCT & "</div>" &
                    "<div class='col-xs-4 col-lg-4'>Valor Log10 (N): " & Math.Round(Math.Log10(lCT), 4) & "</div></div>"
                sRTF &= "<div class='table-responsive'><table class='table table-bordered table-hover'><thead><tr class='info'><th>Espécie</th><th>Número de Indivíduos</th></tr></thead><tbody>"


                For iCelValue = 0 To lrow - 1
                    If oDGV.Rows(iCelValue).Cells(iC).Value <> String.Empty Then _
                    lCel = CDbl(oDGV.Rows(iCelValue).Cells(iC).Value)
                    If lCel > 0 Then
                        sp += 1
                    End If
                    sRTF &= "<tr>"
                    sRTF &= "<td>" & oDGV.Rows(iCelValue).HeaderCell.Value & "</td>"
                    sRTF &= "<td class='text-center'>" & lCel & "</td></tr>"
                Next iCelValue
                lDivMargalef = (sp - 1) / CDbl(Math.Log10(lCT))
                lDivMargalefTotal += lDivMargalef
                sRTF &= "</tbody><tfooter><tr class='text-info'>"
                sRTF &= "<td>Número de Espécies do Levantamento: " & sp & "</td>"
                sRTF &= "<td>Índice Diversidade de Margalef (ɑ): " & Format(Math.Abs(lDivMargalef), "###,###,###,##0.0###") & "</td></tr></tfooter>"
                sRTF &= "</table></div><hr/>"
                sp = 0
            Next iC
            sResult = sHTMLStart & "<h3 class='alert alert-success'>Relatório da Diversidade de Margalef</h3>" & sRTF & sFooter & sHTMLEnd
            Return sResult

        Catch ex As Exception
            LibError.ErrorView(ex)
            Return retError(sHTMLStart, ex, sHTMLEnd)
        End Try
    End Function

    Public Function divGleason(ByRef oDGV As DataGridView)
        Dim lcol As Long, lrow As Long
        Dim iR As Long, iC As Integer
        Dim lCT As Long, lCel As Long
        Dim iCelValue As Integer
        Dim lDivGleason As Double, sLev As String
        Dim sp As Double
        Dim sRTF As String = ""
        lcol = oDGV.ColumnCount
        lrow = oDGV.RowCount
        Try
            If Auditoria(oDGV) = False Then
                If MsgAuditoria() = False Then
                    Return sRetAuditoria

                End If
            End If

            For iC = 0 To lcol - 1 Step 1
                lCT = 0
                For iR = 0 To lrow - 1 Step 1
                    If oDGV.Rows(iR).Cells(iC).Value <> String.Empty Then _
                    lCT += CDbl(oDGV.Rows(iR).Cells(iC).Value)
                Next iR

                sLev = oDGV.Columns(iC).HeaderText.ToString
                sRTF &= "<div class='row'><div class='col-xs-4 col-lg-4'>Levantamento: " & sLev & "</div>"
                sRTF &= "<div class='col-xs-4 col-lg-4'>Valor N: " & lCT & "</div>" &
                    "<div class='col-xs-4 col-lg-4'>Valor Log10 (N): " & Math.Round(Math.Log10(lCT), 4) & "</div></div>"
                sRTF &= "<div class='table-responsive'><table class='table table-bordered table-hover'><thead><tr class='info'><th>Espécie</th><th>Número de Indivíduos</th></tr></thead><tbody>"


                For iCelValue = 0 To lrow - 1
                    If oDGV.Rows(iCelValue).Cells(iC).Value <> String.Empty Then _
                    lCel = CDbl(oDGV.Rows(iCelValue).Cells(iC).Value)
                    If lCel > 0 Then
                        sp += 1
                    End If
                    sRTF &= "<tr>"
                    sRTF &= "<td>" & oDGV.Rows(iCelValue).HeaderCell.Value & "</td>"
                    sRTF &= "<td class='text-center'>" & lCel & "</td></tr>"

                Next iCelValue
                lDivGleason = sp / CDbl(Math.Log10(lCT))
              
                sRTF &= "</tbody><tfooter><tr class='text-info'>"
                sRTF &= "<td>Número de Espécies do Levantamento: " & sp & "</td>"
                sRTF &= "<td>Índice Diversidade de Gleason: " & Format(Math.Abs(lDivGleason), "###,###,###,##0.0###") & "</td></tr></tfooter>"
                sRTF &= "</table></div><hr/>"
                sp = 0
            Next iC
            
            sResult = sHTMLStart & "<h3 class='alert alert-success'>Relatório da Diversidade de Gleason</h3>" & sRTF & sFooter & sHTMLEnd
            Return sResult
        Catch ex As Exception
            LibError.ErrorView(ex)
            Return retError(sHTMLStart, ex, sHTMLEnd)
        End Try
    End Function

    Public Function divMenhinick(ByRef oDGV As DataGridView) As String
        Dim lcol As Long, lrow As Long
        Dim iR As Long, iC As Integer
        Dim lCT As Long, lCel As Long
        Dim iCelValue As Integer
        Dim lDivMenhinick As Double, sLev As String
        Dim sp As Double
        Dim sRTF As String = ""
        lcol = oDGV.ColumnCount
        lrow = oDGV.RowCount
        Try
            If Auditoria(oDGV) = False Then
                If MsgAuditoria() = False Then
                    Return sRetAuditoria

                End If
            End If

            For iC = 0 To lcol - 1 Step 1
                lCT = 0
                For iR = 0 To lrow - 1 Step 1
                    If oDGV.Rows(iR).Cells(iC).Value <> String.Empty Then _
                        lCT += CDbl(oDGV.Rows(iR).Cells(iC).Value)
                Next iR

                sLev = oDGV.Columns(iC).HeaderText.ToString
                sRTF &= "<div class='row'><div class='col-xs-4 col-lg-4'>Levantamento: " & sLev & "</div>"
                sRTF &= "<div class='col-xs-4 col-lg-4'>Valor N: " & lCT & "</div>" &
                    "<div class='col-xs-4 col-lg-4'>Valor Raiz Quadrada: " & Math.Round(Math.Sqrt(lCT), 4) & "</div></div>"
                sRTF &= "<div class='table-responsive'><table class='table table-bordered table-hover'><thead><tr class='info'><th>Espécie</th><th>Número de Indivíduos</th></tr></thead><tbody>"


                For iCelValue = 0 To lrow - 1
                    If oDGV.Rows(iCelValue).Cells(iC).Value <> String.Empty Then _
                    lCel = CDbl(oDGV.Rows(iCelValue).Cells(iC).Value)
                    If lCel > 0 Then
                        sp += 1
                    End If
                    sRTF &= "<tr>"
                    sRTF &= "<td>" & oDGV.Rows(iCelValue).HeaderCell.Value & "</td>"
                    sRTF &= "<td class='text-center'>" & lCel & "</td></tr>"

                Next iCelValue
                lDivMenhinick = sp / Math.Sqrt(lCT)
                sRTF &= "</tbody><tfooter><tr class='text-info'>"
                sRTF &= "<td>Número de Espécies do Levantamento: " & sp & "</td>"
                sRTF &= "<td>Índice Diversidade de Menhinick: " & Format(Math.Abs(lDivMenhinick), "###,###,###,##0.0###") & "</td></tr></tfooter>"
                sRTF &= "</table></div><hr/>"
                sp = 0
            Next iC
            sResult = sHTMLStart & "<h3 class='alert alert-success'>Relatório da Diversidade de Menhinick</h3>" & sRTF & sFooter & sHTMLEnd
            Return sResult


        Catch ex As Exception
            LibError.ErrorView(ex)
            Return retError(sHTMLStart, ex, sHTMLEnd)
        End Try
    End Function

    Public Function divMcIntosh(ByRef oDGV As DataGridView)
        Dim lcol As Long, lrow As Long
        Dim iR As Long, iC As Integer
        Dim lCT As Long, lCel As Long
        Dim lCel2 As Double
        Dim iCelValue As Integer
        Dim lDivMcIntosh As Double, sLev As String
        Dim dU As Double
        Dim sRTF As String = ""
        lcol = oDGV.ColumnCount
        lrow = oDGV.RowCount
        Try

            For iC = 0 To lcol - 1 Step 1
                lCT = 0
                For iR = 0 To lrow - 1 Step 1
                    If oDGV.Rows(iR).Cells(iC).Value <> String.Empty Then _
                    lCT += CDbl(oDGV.Rows(iR).Cells(iC).Value)

                Next iR

                sLev = oDGV.Columns(iC).HeaderText.ToString
                sRTF &= "<div class='row'><div class='col-xs-6 col-lg-6'>Levantamento: " & sLev & "</div>"
                sRTF &= "<div class='col-xs-6 col-lg-6'>Valor N: " & lCT & "</div></div>"
                sRTF &= "<div class='table-responsive'><table class='table table-bordered table-hover'><thead><tr class='info'><th>Espécie</th><th>Número de Indivíduos</th></tr></thead><tbody>"

                For iCelValue = 0 To lrow - 1
                    If oDGV.Rows(iCelValue).Cells(iC).Value <> String.Empty Then _
                    lCel = CDbl(oDGV.Rows(iCelValue).Cells(iC).Value)
                    lCel2 += lCel ^ 2
                    sRTF &= "<tr>"
                    sRTF &= "<td>" & oDGV.Rows(iCelValue).HeaderCell.Value & "</td>"
                    sRTF &= "<td class='text-center'>" & lCel & "</td></tr>"
                Next iCelValue
                dU = Math.Sqrt(lCel2)
                lDivMcIntosh = (lCT - dU) / (lCT - Math.Sqrt(lCT))
                sRTF &= "</tbody><tfooter><tr class='text-info'>"
                sRTF &= "<td>Valor de U: " & Math.Round(dU, 4) & "</td>"
                sRTF &= "<td>Índice Diversidade de McIntosh: " & Format(Math.Abs(lDivMcIntosh), "###,###,###,##0.0###") & "</td></tr></tfooter>"
                sRTF &= "</table></div><hr/>"
                dU = 0
                lCel2 = 0
            Next iC
            sResult = sHTMLStart & "<h3 class='alert alert-success'>Relatório da Diversidade de McIntosh</h3>" & sRTF & sFooter & sHTMLEnd
            Return sResult
        Catch ex As Exception
            LibError.ErrorView(ex)
            Return retError(sHTMLStart, ex, sHTMLEnd)
        End Try
    End Function

    Public Function divBrillouin(ByRef oDGV As DataGridView) As String
        Dim lcol As Long, lrow As Long
        Dim iR As Long, iC As Integer
        Dim lCT As Long, lCel As Long
        Dim iCelValue As Integer
        Dim lDivBrillouin As Double
        Dim sLev As String
        Dim FactN As Double, Factni As Double, FactRet As Double
        Dim sRTF As String = ""

        lcol = oDGV.ColumnCount
        lrow = oDGV.RowCount
        Try
            If Auditoria(oDGV) = False Then
                If MsgAuditoria() = False Then
                    Return sRetAuditoria

                End If
            End If
            iLog = LibLoadConfig.LoadConfig("CONFIG", "LOGBASE", "2")
            For iC = 0 To lcol - 1 Step 1
                lCT = 0
                For iR = 0 To lrow - 1 Step 1
                    If oDGV.Rows(iR).Cells(iC).Value <> String.Empty Then _
                    lCT += oDGV.Rows(iR).Cells(iC).Value
                Next iR
                FactN = lCT
                
                FactN = Math.Log(LibCalculate.Factorial(FactN), iLog)
                
                sLev = oDGV.Columns(iC).HeaderText.ToString
                sRTF &= "<div class='row'><div class='col-xs-4 col-lg-4'>Levantamento: " & sLev & "</div>"
                sRTF &= "<div class='col-xs-4 col-lg-4'>Valor N: " & lCT & "</div>"
                sRTF &= "<div class='col-xs-4 col-lg-4'>Base Logarítmica: " & iLog & "</div></div>"
                sRTF &= "<div class='table-responsive'><table class='table table-bordered table-hover'><thead><tr class='info'><th colspan=2>Espécie</th><th>Número de Indivíduos</th></tr></thead><tbody>"

                For iCelValue = 0 To lrow - 1
                    If oDGV.Rows(iCelValue).Cells(iC).Value <> String.Empty Then
                        lCel = CDbl(oDGV.Rows(iCelValue).Cells(iC).Value)
                    Else
                        lCel = 0
                    End If
                    If lCel > 0 Then
                        
                        FactRet = LibCalculate.Factorial(lCel)
                        Factni += Math.Log(FactRet, iLog)
                       
                        sRTF &= "<tr>"
                        sRTF &= "<td colspan=2>" & oDGV.Rows(iCelValue).HeaderCell.Value & "</td>"
                        sRTF &= "<td class='text-center'>" & lCel & "</td></tr>"
                    End If
                Next iCelValue
                lDivBrillouin = (FactN - Factni) / (lCT)
                If lDivBrillouin = Double.NaN Or IsNumeric(lDivBrillouin) = False Then lDivBrillouin = 0
                sRTF &= "</tbody><tfooter><tr class='text-info'>"
                sRTF &= "<td>Fatorial de Log de N!: " & Math.Round(FactN, 4) & "</td>"
                sRTF &= "<td>Somatório do log n!: " & Math.Round(Factni, 4) & "</td>"
                sRTF &= "<td>Índice Diversidade de Brillouin: " & Format(lDivBrillouin, "###,###,###,##0.0###") & "</td></tr></tfooter>"
                sRTF &= "</table></div><hr/>"

                Factni = 0
                FactN = 0
            Next iC
            sResult = sHTMLStart & "<h3 class='alert alert-success'>Relatório da Diversidade de Brillouin</h3>" & sRTF & sFooter & sHTMLEnd
            Return sResult

        Catch ex As Exception
            LibError.ErrorView(ex)
            Return retError(sHTMLStart, ex, sHTMLEnd)
        End Try
    End Function

    Public Function divTotal(ByRef oDGV As DataGridView) As String
        Dim lcol As Long, lrow As Long
        Dim iR As Long, iC As Integer
        Dim lCT As Long, lCel As Long
        Dim iCelValue As Integer
        Dim lDivTotal As Double, sLev As String
        Dim lFi As Double, lFiMinus As Double, lFixFiMinus As Double, lWi As Double
        Dim sRTF As String = ""
        lcol = oDGV.ColumnCount
        lrow = oDGV.RowCount
        Try
            If Auditoria(oDGV) = False Then
                If MsgAuditoria() = False Then
                    Return sRetAuditoria

                End If
            End If
            For iC = 0 To lcol - 1 Step 1
                lCT = 0
                For iR = 0 To lrow - 1 Step 1
                    If oDGV.Rows(iR).Cells(iC).Value <> String.Empty Then _
                    lCT += oDGV.Rows(iR).Cells(iC).Value
                Next iR
                sLev = oDGV.Columns(iC).HeaderText.ToString
                sRTF &= "<div class='row'><div class='col-xs-6 col-lg-6'>Levantamento: " & sLev & "</div>"
                sRTF &= "<div class='col-xs-6 col-lg-6'>Valor N: " & lCT & "</div></div>"
                sRTF &= "<div class='table-responsive'><table class='table table-bordered table-hover'><thead><tr class='info'><th>Espécie</th><th>Número de Indivíduos</th></tr></thead><tbody>"
                For iCelValue = 0 To lrow - 1
                    If oDGV.Rows(iCelValue).Cells(iC).Value <> String.Empty Then
                        lCel = CDbl(oDGV.Rows(iCelValue).Cells(iC).Value)
                    Else
                        lCel = 0
                    End If
                    If lCel <> 0 Then
                        lFi = lCel / lCT
                        lFiMinus = 1 - lFi
                        lFixFiMinus = lFi * lFiMinus
                        lWi = 1 / lFi
                    End If
                    sRTF &= "<tr>"
                    sRTF &= "<td>" & oDGV.Rows(iCelValue).HeaderCell.Value & "</td>"
                    sRTF &= "<td class='text-center'>" & lCel & "</td></tr>"
                Next iCelValue
                lDivTotal = lWi * lFixFiMinus
                sRTF &= "</tbody><tfooter><tr class='text-info'>"
                sRTF &= "<td>Valor de Wi: " & Math.Round(lWi, 4) & "</td>"
                sRTF &= "<td>Índice Diversidade de Total:  " & Format(lDivTotal, "###,###,###,##0.0###") & "</td></tr></tfooter>"
                sRTF &= "</table></div><hr/>"

                lFi = 0
                lFiMinus = 0
                lFixFiMinus = 0
            Next iC

            sResult = sHTMLStart & "<h3 class='alert alert-success'>Relatório da Diversidade de Total</h3>" & sRTF & sFooter & sHTMLEnd
            Return sResult
        Catch ex As Exception
            LibError.ErrorView(ex)
            Return retError(sHTMLStart, ex, sHTMLEnd)
        End Try
    End Function

End Class
