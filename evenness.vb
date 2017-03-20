Imports System.Windows.Forms

Public Class evenness
    Private libCalculate As New calculate
    Sub New()
        Main()
    End Sub

    Public Function Dmax(ByVal ODGV As DataGridView) As String
        Dim lcol As Long, lrow As Long
        Dim iR As Long, iC As Integer
        Dim lCT As Long, lCel As Long
        Dim iCelValue As Integer
        Dim lFi As Double
        Dim sLev As String, somaT As Double
        Dim lNvsN As Double, dDivMax As Double
        Dim sRTF As String = ""
        Dim nSp As Double
        lcol = ODGV.ColumnCount
        lrow = ODGV.RowCount
        Try
            If Auditoria(ODGV) = False Then
                If MsgAuditoria() = False Then
                    Return sRetAuditoria
                    Exit Function
                End If
            End If

            For iC = 0 To lcol - 1 Step 1
                lCT = 0
                For iR = 0 To lrow - 1 Step 1
                    If ODGV.Rows(iR).Cells(iC).Value <> String.Empty Then _
                    lCT += CDbl(ODGV.Rows(iR).Cells(iC).Value)
                Next iR

                sLev = ODGV.Columns(iC).HeaderText.ToString
                sRTF &= "<div class='row'><div class='col-xs-6 col-lg-6'>Levantamento: " & sLev & "</div>"
                sRTF &= "<div class='col-xs-6 col-lg-6'>Valor N: " & lCT & "</div></div>"
                sRTF &= "<div class='table-responsive'><table class='table table-bordered table-hover' ><thead><tr class='info'><th>Espécie</th><th>Número de Indivíduos</th></tr></thead><tbody>"
                lNvsN = lCT * (lCT - 1)
                For iCelValue = 0 To lrow - 1
                    If ODGV.Rows(iCelValue).Cells(iC).Value <> String.Empty Then
                        lCel = CDbl(ODGV.Rows(iCelValue).Cells(iC).Value)
                        sRTF &= "<tr>"
                        sRTF &= "<td>" & ODGV.Rows(iCelValue).HeaderCell.Value & "</td>"
                        sRTF &= "<td class='text-center'>" & lCel & "</td>"

                        sRTF &= "</tr>"
                    End If
                    If lCel > 0 Then
                        nSp += 1
                    End If

                Next iCelValue
                dDivMax = ((nSp - 1) / (nSp)) * (lCT / (lCT - 1))

                sRTF &= "</tbody><tfooter><tr class='text-info'>"
                sRTF &= "<td colspan=2>Regularidade D<sub>max</sub>: " & Format(Math.Abs(dDivMax), "###,###,###,##0.0###") & "</td>"
                sRTF &= "</tr></tfooter>"
                sRTF &= "</table></div><hr/>"
                somaT = 0
                lFi = 0
                nSp = 0
            Next iC


            sResult = sHTMLStart & "<h3 class='alert alert-success'>Regularidade D<sub>max</sub></h3>" & sRTF & sFooter & sHTMLEnd
            Return sResult

        Catch ex As Exception

            MessageBox.Show("Erro: " & ex.Message & vbNewLine & vbNewLine & "Local de Origem: " & ex.StackTrace, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Return retError(sHTMLStart, ex, sHTMLEnd)
        End Try
    End Function

    Public Function dminusmax(ByVal ODGV As DataGridView) As String
        Dim lcol As Long, lrow As Long
        Dim iR As Long, iC As Integer
        Dim lCT As Long, lCel As Long
        Dim iCelValue As Integer
        Dim lFi As Double
        Dim sLev As String, somaT As Double
        Dim lNvsN As Double, dDivMinusMax As Double
        Dim sRTF As String = ""
        Dim nSp As Double
        lcol = ODGV.ColumnCount
        lrow = ODGV.RowCount
        Try
            If Auditoria(ODGV) = False Then
                If MsgAuditoria() = False Then
                    Return sRetAuditoria
                    Exit Function
                End If
            End If

            For iC = 0 To lcol - 1 Step 1
                lCT = 0
                For iR = 0 To lrow - 1 Step 1
                    If ODGV.Rows(iR).Cells(iC).Value <> String.Empty Then _
                    lCT += CDbl(ODGV.Rows(iR).Cells(iC).Value)
                Next iR

                sLev = ODGV.Columns(iC).HeaderText.ToString
                sRTF &= "<div class='row'><div class='col-xs-6 col-lg-6'>Levantamento: " & sLev & "</div>"
                sRTF &= "<div class='col-xs-6 col-lg-6'>Valor N: " & lCT & "</div></div>"
                sRTF &= "<div class='table-responsive'><table class='table table-bordered table-hover' ><thead><tr class='info'><th>Espécie</th><th>Número de Indivíduos</th></tr></thead><tbody>"
                lNvsN = lCT * (lCT - 1)
                For iCelValue = 0 To lrow - 1
                    If ODGV.Rows(iCelValue).Cells(iC).Value <> String.Empty Then
                        lCel = CDbl(ODGV.Rows(iCelValue).Cells(iC).Value)
                        sRTF &= "<tr>"
                        sRTF &= "<td>" & ODGV.Rows(iCelValue).HeaderCell.Value & "</td>"
                        sRTF &= "<td class='text-center'>" & lCel & "</td>"

                        sRTF &= "</tr>"
                    End If
                    If lCel > 0 Then
                        nSp += 1
                    End If

                Next iCelValue
                dDivMinusMax = nSp * ((lCT - 1) / (lCT - nSp))
                sRTF &= "</tbody><tfooter><tr class='text-info'>"
                sRTF &= "<td colspan=2>Regularidade <em>d<sub>max</sub></em>: " & Format(Math.Abs(dDivMinusMax), "###,###,###,##0.0###") & "</td>"
                sRTF &= "</tr></tfooter>"
                sRTF &= "</table></div><hr/>"
                somaT = 0
                lFi = 0
                nSp = 0
            Next iC
            'sRTF &= "Índice Diversidade de Simpson Total: " & Format(Math.Abs(lDivSimTotal), "###,###,###,##0.0###") & vbNewLin

            sResult = sHTMLStart & "<h3 class='alert alert-success'>Regularidade <em>d<sub>max</sub></em></h3>" & sRTF & sFooter & sHTMLEnd
            Return sResult

        Catch ex As Exception

            MessageBox.Show("Erro: " & ex.Message & vbNewLine & vbNewLine & "Local de Origem: " & ex.StackTrace, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Return retError(sHTMLStart, ex, sHTMLEnd)
        End Try
    End Function

    Public Function deltamax(ByVal ODGV As DataGridView) As String
        Dim lcol As Long, lrow As Long
        Dim iR As Long, iC As Integer
        Dim lCT As Long, lCel As Long
        Dim iCelValue As Integer
        Dim lFi As Double
        Dim sLev As String, somaT As Double
        Dim lNvsN As Double, dDeltaMax As Double
        Dim sRTF As String = ""
        Dim nSp As Double
        lcol = ODGV.ColumnCount
        lrow = ODGV.RowCount
        Try
            If Auditoria(ODGV) = False Then
                If MsgAuditoria() = False Then
                    Return sRetAuditoria
                    Exit Function
                End If
            End If

            For iC = 0 To lcol - 1 Step 1
                lCT = 0
                For iR = 0 To lrow - 1 Step 1
                    If ODGV.Rows(iR).Cells(iC).Value <> String.Empty Then _
                    lCT += CDbl(ODGV.Rows(iR).Cells(iC).Value)
                Next iR

                sLev = ODGV.Columns(iC).HeaderText.ToString
                sRTF &= "<div class='row'><div class='col-xs-6 col-lg-6'>Levantamento: " & sLev & "</div>"
                sRTF &= "<div class='col-xs-6 col-lg-6'>Valor N: " & lCT & "</div></div>"
                sRTF &= "<div class='table-responsive'><table class='table table-bordered table-hover' ><thead><tr class='info'><th>Espécie</th><th>Número de Indivíduos</th></tr></thead><tbody>"
                lNvsN = lCT * (lCT - 1)
                For iCelValue = 0 To lrow - 1
                    If ODGV.Rows(iCelValue).Cells(iC).Value <> String.Empty Then
                        lCel = CDbl(ODGV.Rows(iCelValue).Cells(iC).Value)
                        sRTF &= "<tr>"
                        sRTF &= "<td>" & ODGV.Rows(iCelValue).HeaderCell.Value & "</td>"
                        sRTF &= "<td class='text-center'>" & lCel & "</td>"

                        sRTF &= "</tr>"
                    End If
                    If lCel > 0 Then
                        nSp += 1
                    End If

                Next iCelValue
                dDeltaMax = 1 - (1 / nSp)

                sRTF &= "</tbody><tfooter><tr class='text-info'>"
                sRTF &= "<td colspan=2>Regularidade Δ<sub>max</sub>: " & Format(Math.Abs(dDeltaMax), "###,###,###,##0.0###") & "</td>"
                sRTF &= "</tr></tfooter>"
                sRTF &= "</table></div><hr/>"
                somaT = 0
                lFi = 0
                nSp = 0
            Next iC
            'sRTF &= "Índice Diversidade de Simpson Total: " & Format(Math.Abs(lDivSimTotal), "###,###,###,##0.0###") & vbNewLin

            sResult = sHTMLStart & "<h3 class='alert alert-success'>Regularidade Δ<sub>max</sub></h3>" & sRTF & sFooter & sHTMLEnd
            Return sResult

        Catch ex As Exception

            MessageBox.Show("Erro: " & ex.Message & vbNewLine & vbNewLine & "Local de Origem: " & ex.StackTrace, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Return retError(sHTMLStart, ex, sHTMLEnd)
        End Try
    End Function

    Public Function Hmax(ByVal ODGV As DataGridView) As String
        Dim lcol As Long, lrow As Long
        Dim iR As Long, iC As Integer
        Dim lCT As Long, lCel As Long
        Dim iCelValue As Integer
        Dim lFi As Double, dRemainde As Double, iCValue As Integer, dLogNFact As Double, dLogCplus1 As Double, dlogC As Double
        Dim sLev As String, somaT As Double
        Dim lNvsN As Double, dHMax As Double
        Dim sRTF As String = ""
        Dim nSp As Double
        lcol = ODGV.ColumnCount
        lrow = ODGV.RowCount
        Try
            If Auditoria(ODGV) = False Then
                If MsgAuditoria() = False Then
                    Return sRetAuditoria
                    Exit Function
                End If
            End If
            iLog = LibConfig.LoadConfig("CONFIG", "LOGBASE", "2")
            For iC = 0 To lcol - 1 Step 1
                lCT = 0
                For iR = 0 To lrow - 1 Step 1
                    If ODGV.Rows(iR).Cells(iC).Value <> String.Empty Then _
                    lCT += CDbl(ODGV.Rows(iR).Cells(iC).Value)
                Next iR

                sLev = ODGV.Columns(iC).HeaderText.ToString
                sRTF &= "<div class='row'><div class='col-xs-6 col-lg-6'>Levantamento: " & sLev & "</div>"
                sRTF &= "<div class='col-xs-6 col-lg-6'>Valor N: " & lCT & "</div></div>"
                sRTF &= "<div class='table-responsive'><table class='table table-bordered table-hover' ><thead><tr class='info'><th>Espécie</th><th>Número de Indivíduos</th></tr></thead><tbody>"
                lNvsN = lCT * (lCT - 1)
                For iCelValue = 0 To lrow - 1
                    If ODGV.Rows(iCelValue).Cells(iC).Value <> String.Empty Then
                        lCel = CDbl(ODGV.Rows(iCelValue).Cells(iC).Value)
                        sRTF &= "<tr>"
                        sRTF &= "<td>" & ODGV.Rows(iCelValue).HeaderCell.Value & "</td>"
                        sRTF &= "<td class='text-center'>" & lCel & "</td>"

                        sRTF &= "</tr>"
                    End If
                    If lCel > 0 Then
                        nSp += 1
                    End If

                Next iCelValue
                dRemainde = lCT Mod nSp
                iCValue = Math.Abs(lCT / nSp)
                dLogNFact = Math.Log(libCalculate.Factorial(lCT), iLog)
                dlogC = Math.Log(iCValue, iLog)
                dLogCplus1 = Math.Log(libCalculate.Factorial(iCValue + 1), iLog)
                dHMax = (dLogNFact - (nSp - dRemainde) * (dlogC - dRemainde * dLogCplus1)) / lCT
                sRTF &= "</tbody><tfooter><tr class='text-info'>"
                sRTF &= "<td colspan=2>Regularidade H<sub>max</sub>: " & Format(Math.Abs(dHMax), "###,###,###,##0.0###") & "</td>"
                sRTF &= "</tr></tfooter>"
                sRTF &= "</table></div><hr/>"
                somaT = 0
                lFi = 0
                nSp = 0
            Next iC
            'sRTF &= "Índice Diversidade de Simpson Total: " & Format(Math.Abs(lDivSimTotal), "###,###,###,##0.0###") & vbNewLin

            sResult = sHTMLStart & "<h3 class='alert alert-success'>Regularidade H<sub>max</sub></h3>" & sRTF & sFooter & sHTMLEnd
            Return sResult

        Catch ex As Exception

            MessageBox.Show("Erro: " & ex.Message & vbNewLine & vbNewLine & "Local de Origem: " & ex.StackTrace, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Return retError(sHTMLStart, ex, sHTMLEnd)
        End Try
    End Function

End Class
