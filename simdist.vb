Imports System.Windows.Forms
'' Jaccard, Euclidiana e Bray-Curtis - 
'' http://ecovirtual.ib.usp.br/doku.php?id=ecovirt:roteiro:comuni:comuni_classr&do=
'' http://prf.osu.cz/kbe/dokumenty/sw/ComEcoPaC/ComEcoPaC.pdf
'' Manhattan -  http://cc.oulu.fi/~jarioksa/softhelp/vegan/html/vegdist.html
'' Morisita -
'' http://cc.oulu.fi/~jarioksa/softhelp/vegan/html/dispindmorisita.html 
'' http://www.zoology.ubc.ca/~krebs/downloads/krebs_chapter_12_2014.pdf
'' http://www.statisticshowto.com/morisita-index/
'' Morisita-Horn- https://en.wikipedia.org/wiki/Morisita%27s_overlap_index
'' Sorenson Similarity - C:\Users\WilliamCosta\Documents\DivEs - Documentação\ALS 5932_Lecture Richness and Diversity 10_08_08.pdf
'' Canberra - http://www.int-res.com/articles/meps/5/m005p125.pdf
'' Renkonen index - http://prf.osu.cz/kbe/dokumenty/sw/ComEcoPaC/ComEcoPaC.pdf
Public Class simdist
    Private LibCalculate As New decalculate.calculate
   ' Private LibError As New deconfig.ErrorManager
    Sub New()
        Main()
    End Sub

    Public Function Jaccard(ByVal ODGV As DataGridView, ByVal iColumn1 As Integer, ByVal iColumn2 As Integer) As String
        Dim lRow As Integer = ODGV.RowCount
        Dim iR As Integer
        Dim nSp1 As Integer, nSp2 As Integer, nSpCom As Integer
        Dim sRTF As String = ""
        Dim dJaccard As Double
        Try

            If Auditoria(ODGV) = False Then
                If MsgAuditoria() = False Then
                    Return sRetAuditoria
                    Exit Function
                End If
            End If
            For iR = 0 To lRow - 1
                If ODGV.Rows(iR).Cells(iColumn1).Value > 0 Then
                    nSp1 += 1
                End If
                If ODGV.Rows(iR).Cells(iColumn2).Value > 0 Then
                    nSp2 += 1
                End If
                If ODGV.Rows(iR).Cells(iColumn1).Value > 0 And ODGV.Rows(iR).Cells(iColumn2).Value > 0 Then
                    nSpCom += 1
                End If
            Next
            dJaccard = (nSpCom) / (nSp1 + nSp2 - nSpCom)
            sRTF &= "<ul class='list-group'>"
            sRTF &= "<li class='list-group-item'>Número de espécies da amostra " & ODGV.Columns(iColumn1).HeaderCell.Value & ": " & nSp1 & "</li>"
            sRTF &= "<li class='list-group-item'>Número de espécies da amostra " & ODGV.Columns(iColumn2).HeaderCell.Value & ": " & nSp2 & "</li>"
            sRTF &= "<li class='list-group-item'>Número de espécies em comum: " & nSpCom & "</li>"
            sRTF &= "<li class='list-group-item'>Índice de Similaridade de Jaccard: " & Math.Round(dJaccard, 4) & "</li>"
            sRTF &= "</ul>"
            sResult = sHTMLStart & "<h3 class='alert alert-success'>Similaridade de Jaccard</h3>" & sRTF & sFooter & sHTMLEnd
            Return sResult

        Catch ex As Exception
            LibError.ErrorView(ex)
            Return retError(sHTMLStart, ex, sHTMLEnd)
        End Try
    End Function

    Public Function Euclidiana(ByVal ODGV As DataGridView, ByVal iColumn1 As Integer, ByVal iColumn2 As Integer) As String
        Dim lRow As Integer = ODGV.RowCount
        Dim iR As Integer
        Dim nSp1 As Integer, nSp2 As Integer, dn1n2 As Double
        Dim sRTF As String = ""
        Dim dEuclid As Double
        Try

            If Auditoria(ODGV) = False Then
                If MsgAuditoria() = False Then
                    Return sRetAuditoria
                    Exit Function
                End If
            End If
            For iR = 0 To lRow - 1
                If ODGV.Rows(iR).Cells(iColumn1).Value > 0 Then
                    nSp1 += ODGV.Rows(iR).Cells(iColumn1).Value
                End If
                If ODGV.Rows(iR).Cells(iColumn2).Value > 0 Then
                    nSp2 += ODGV.Rows(iR).Cells(iColumn2).Value
                End If

                dn1n2 += (ODGV.Rows(iR).Cells(iColumn1).Value - ODGV.Rows(iR).Cells(iColumn2).Value) ^ 2
            Next

            dEuclid = Math.Sqrt(dn1n2)
            sRTF &= "<ul class='list-group'>"
            sRTF &= "<li class='list-group-item'>Número de indivíduos da amostra " & ODGV.Columns(iColumn1).HeaderCell.Value & ": " & nSp1 & "</li>"
            sRTF &= "<li class='list-group-item'>Número de indivíduos da amostra " & ODGV.Columns(iColumn2).HeaderCell.Value & ": " & nSp2 & "</li>"
            sRTF &= "<li class='list-group-item'>Distância Euclidiana: " & Math.Round(dEuclid, 4) & "</li>"
            sRTF &= "</ul>"
            sResult = sHTMLStart & "<h3 class='alert alert-success'>Distância Euclidiana</h3>" & sRTF & sFooter & sHTMLEnd
            Return sResult
        Catch ex As Exception
            LibError.ErrorView(ex)
            Return retError(sHTMLStart, ex, sHTMLEnd)
        End Try
    End Function

    Public Function BCSim(ByVal ODGV As DataGridView, ByVal iColumn1 As Integer, ByVal iColumn2 As Integer) As String
        Dim lRow As Integer = ODGV.RowCount
        Dim iR As Integer, nSp1 As Integer, nSp2 As Integer
        Dim iNum As Integer, dMin As Double, dSum As Double
        Dim sRTF As String = ""
        Dim dSimiliar As Double
        Try
            If Auditoria(ODGV) = False Then
                If MsgAuditoria() = False Then
                    Return sRetAuditoria
                    Exit Function
                End If
            End If
            For iR = 0 To lRow - 1
                If ODGV.Rows(iR).Cells(iColumn1).Value > 0 Then
                    nSp1 += ODGV.Rows(iR).Cells(iColumn1).Value
                    iNum += ODGV.Rows(iR).Cells(iColumn1).Value
                End If
                If ODGV.Rows(iR).Cells(iColumn2).Value > 0 Then
                    nSp2 += ODGV.Rows(iR).Cells(iColumn2).Value
                    iNum += ODGV.Rows(iR).Cells(iColumn2).Value
                End If

            Next
            For iR = 0 To lRow - 1

                dMin = Math.Min(CByte(ODGV.Rows(iR).Cells(iColumn1).Value), CByte(ODGV.Rows(iR).Cells(iColumn2).Value))
                dSum += dMin
            Next
            dSimiliar = (2 * dSum) / iNum
            sRTF &= "<ul class='list-group'>"
            sRTF &= "<li class='list-group-item'>Número de indivíduos da amostra " & ODGV.Columns(iColumn1).HeaderCell.Value & ": " & nSp1 & "</li>"
            sRTF &= "<li class='list-group-item'>Número de indivíduos da amostra " & ODGV.Columns(iColumn2).HeaderCell.Value & ": " & nSp2 & "</li>"
            sRTF &= "<li class='list-group-item'>Índice de Similaridade Bray-Curtis: " & Math.Round(dSimiliar, 4) & "</li>"
            sRTF &= "<li class='list-group-item'>Índice de Disimilaridade Bray-Curtis: " & Math.Round(1 - dSimiliar, 4) & "</li>"
            sRTF &= "</ul>"
            sResult = sHTMLStart & "<h3 class='alert alert-success'>Similaridade de Bray-Curtis</h3>" & sRTF & sFooter & sHTMLEnd
            Return sResult
        Catch ex As Exception
            LibError.ErrorView(ex)
            Return retError(sHTMLStart, ex, sHTMLEnd)
        End Try
    End Function

    Public Function BCDist(ByVal ODGV As DataGridView, ByVal iColumn1 As Integer, ByVal iColumn2 As Integer) As String
        Dim lRow As Integer = ODGV.RowCount
        Dim iR As Integer
        Dim nSp1 As Integer, nSp2 As Integer, dn1n2 As Double
        Dim sRTF As String = ""
        Dim dDist As Double
        Try
            If Auditoria(ODGV) = False Then
                If MsgAuditoria() = False Then
                    Return sRetAuditoria
                    Exit Function
                End If
            End If
            For iR = 0 To lRow - 1
                If ODGV.Rows(iR).Cells(iColumn1).Value > 0 Then
                    nSp1 += ODGV.Rows(iR).Cells(iColumn1).Value

                End If
                If ODGV.Rows(iR).Cells(iColumn2).Value > 0 Then
                    nSp2 += ODGV.Rows(iR).Cells(iColumn2).Value
                End If
                dn1n2 += Math.Abs(ODGV.Rows(iR).Cells(iColumn1).Value - ODGV.Rows(iR).Cells(iColumn2).Value)
            Next
            dDist = dn1n2 / (nSp1 + nSp1)
            sRTF &= "<ul class='list-group'>"
            sRTF &= "<li class='list-group-item'>Número de indivíduos da amostra " & ODGV.Columns(iColumn1).HeaderCell.Value & ": " & nSp1 & "</li>"
            sRTF &= "<li class='list-group-item'>Número de indivíduos da amostra " & ODGV.Columns(iColumn2).HeaderCell.Value & ": " & nSp2 & "</li>"
            sRTF &= "<li class='list-group-item'>Distância de Bray-Curtis: " & Math.Round(dDist, 4) & "</li>"
            sRTF &= "</ul>"
            sResult = sHTMLStart & "<h3 class='alert alert-success'>Distância de Bray-Curtis</h3>" & sRTF & sFooter & sHTMLEnd
            Return sResult
        Catch ex As Exception
            LibError.ErrorView(ex)
            Return retError(sHTMLStart, ex, sHTMLEnd)
        End Try
    End Function

    Public Function Manhattan(ByVal ODGV As DataGridView, ByVal iColumn1 As Integer, ByVal iColumn2 As Integer) As String
        Dim lRow As Integer = ODGV.RowCount
        Dim iR As Integer
        Dim nSp1 As Integer, nSp2 As Integer, dn1n2 As Double
        Dim sRTF As String = ""
        Dim dDist As Double
        Try
            If Auditoria(ODGV) = False Then
                If MsgAuditoria() = False Then
                    Return sRetAuditoria
                    Exit Function
                End If
            End If
            For iR = 0 To lRow - 1
                If ODGV.Rows(iR).Cells(iColumn1).Value > 0 Then
                    nSp1 += ODGV.Rows(iR).Cells(iColumn1).Value

                End If
                If ODGV.Rows(iR).Cells(iColumn2).Value > 0 Then
                    nSp2 += ODGV.Rows(iR).Cells(iColumn2).Value
                End If
                dn1n2 += Math.Abs(ODGV.Rows(iR).Cells(iColumn1).Value - ODGV.Rows(iR).Cells(iColumn2).Value)
            Next
            dDist = dn1n2
            sRTF &= "<ul class='list-group'>"
            sRTF &= "<li class='list-group-item'>Número de indivíduos da amostra " & ODGV.Columns(iColumn1).HeaderCell.Value & ": " & nSp1 & "</li>"
            sRTF &= "<li class='list-group-item'>Número de indivíduos da amostra " & ODGV.Columns(iColumn2).HeaderCell.Value & ": " & nSp2 & "</li>"
            sRTF &= "<li class='list-group-item'>Distância de Manhattan: " & Math.Round(dDist, 4) & "</li>"
            sRTF &= "</ul>"
            sResult = sHTMLStart & "<h3 class='alert alert-success'>Distância de Manhattan</h3>" & sRTF & sFooter & sHTMLEnd
            Return sResult
        Catch ex As Exception
            LibError.ErrorView(ex)
            Return retError(sHTMLStart, ex, sHTMLEnd)
        End Try
    End Function

    Public Function Morisita(ByVal ODGV As DataGridView, ByVal iColumn1 As Integer, ByVal iColumn2 As Integer) As String

        Dim lRow As Integer = ODGV.RowCount
        Dim iR As Integer
        Dim nSp1 As Integer, nSp2 As Integer, dn1n2 As Double
        Dim sRTF As String = ""
        Dim dSimilar As Double, dDisp1 As Double, dDisp2 As Double
        Dim dSumX1Doub As Double, dLambda1 As Double
        Dim dSumX2Doub As Double, dLambda2 As Double
        Dim dSumLamb1 As Double, dSumLamb2 As Double ' soma numeradores
        Dim dSumDenLamb1 As Double, dSumDenLamb2 As Double 'soma Denominadores
        Dim dMor1 As Double
        Dim dValue As String
        Try
            If Auditoria(ODGV) = False Then
                If MsgAuditoria() = False Then
                    Return sRetAuditoria
                    Exit Function
                End If
            End If
            For iR = 0 To lRow - 1
                dValue = ODGV.Rows(iR).Cells(iColumn1).Value
                nSp1 += dValue
                dSumX1Doub += dValue ^ 2
                If dValue > 1 Then
                    dSumLamb1 += dValue * (dValue - 1)
                    dSumDenLamb1 += dValue - 1
                End If

                dValue = ODGV.Rows(iR).Cells(iColumn2).Value
                nSp2 += dValue
                dSumX2Doub += dValue ^ 2
                If dValue > 1 Then
                    dSumLamb2 += dValue * (dValue - 1)
                    dSumDenLamb2 += dValue - 1
                End If
                dn1n2 += ODGV.Rows(iR).Cells(iColumn1).Value * ODGV.Rows(iR).Cells(iColumn2).Value
            Next
            dLambda1 = dSumLamb1 / (nSp1 * dSumDenLamb1)
            dLambda2 = dSumLamb2 / (nSp2 * dSumDenLamb2)
            dMor1 = (2 * dn1n2) / ((dLambda1 + dLambda2) * (nSp1 * nSp2))
            dDisp1 = lRow * (dSumX1Doub - nSp1) / ((nSp1 ^ 2) - nSp1)
            dDisp2 = lRow * (dSumX2Doub - nSp2) / ((nSp2 ^ 2) - nSp2)
            dSimilar = dn1n2
            sRTF &= "<ul class='list-group'>"
            sRTF &= "<li class='list-group-item'>Número de indivíduos da amostra " & ODGV.Columns(iColumn1).HeaderCell.Value & ": " & nSp1 & "</li>"
            sRTF &= "<li class='list-group-item'>Número de indivíduos da amostra " & ODGV.Columns(iColumn2).HeaderCell.Value & ": " & nSp2 & "</li>"
            sRTF &= "<li class='list-group-item'>Similaridade de Morisita: " & Math.Round(dMor1, 4) & "</li>"
            sRTF &= "<li class='list-group-item'>Dispersão de Morisita amostra " & ODGV.Columns(iColumn1).HeaderCell.Value & ": " & Math.Round(dDisp1, 4) & "</li>"
            sRTF &= "<li class='list-group-item'>Dispersão de Morisita amostra " & ODGV.Columns(iColumn2).HeaderCell.Value & ": " & Math.Round(dDisp2, 4) & "</li>"
            sRTF &= "</ul>"
            sResult = sHTMLStart & "<h3 class='alert alert-success'>Índice de Morisita</h3>" & sRTF & sFooter & sHTMLEnd
            Return sResult
        Catch ex As Exception
            LibError.ErrorView(ex)
            Return retError(sHTMLStart, ex, sHTMLEnd)
        End Try
    End Function

    Public Function MorisitaHorn(ByVal ODGV As DataGridView, ByVal iColumn1 As Integer, ByVal iColumn2 As Integer) As String
        Dim lRow As Integer = ODGV.RowCount
        Dim iR As Integer
        Dim nSp1 As Integer, nSp2 As Integer, dn1n2 As Double
        Dim sRTF As String = ""
        Dim dSimilar As Double
        Dim dSumX1Doub As Double, dLambda1 As Double
        Dim dSumX2Doub As Double, dLambda2 As Double
        Dim dMorHorn As Double
        Dim dValue As String
        Try
            If Auditoria(ODGV) = False Then
                If MsgAuditoria() = False Then
                    Return sRetAuditoria
                    Exit Function
                End If
            End If
            For iR = 0 To lRow - 1
                dValue = ODGV.Rows(iR).Cells(iColumn1).Value
                nSp1 += dValue
                dSumX1Doub += dValue ^ 2

                dValue = ODGV.Rows(iR).Cells(iColumn2).Value
                nSp2 += dValue
                dSumX2Doub += dValue ^ 2

                dn1n2 += (ODGV.Rows(iR).Cells(iColumn1).Value * ODGV.Rows(iR).Cells(iColumn2).Value)
            Next
            dLambda1 = dSumX1Doub / (nSp1 ^ 2)
            dLambda2 = dSumX2Doub / (nSp2 ^ 2)
            dMorHorn = (2 * dn1n2) / ((dLambda1 + dLambda2) * (nSp1 * nSp2))

            dSimilar = dn1n2
            sRTF &= "<ul class='list-group'>"
            sRTF &= "<li class='list-group-item'>Número de indivíduos da amostra " & ODGV.Columns(iColumn1).HeaderCell.Value & ": " & nSp1 & "</li>"
            sRTF &= "<li class='list-group-item'>Número de indivíduos da amostra " & ODGV.Columns(iColumn2).HeaderCell.Value & ": " & nSp2 & "</li>"
            sRTF &= "<li class='list-group-item'>Similaridade de Morisita-Horn: " & Math.Round(dMorHorn, 4) & "</li>"
            sRTF &= "</ul>"
            sResult = sHTMLStart & "<h3 class='alert alert-success'>Índice de Morisita-Horn</h3>" & sRTF & sFooter & sHTMLEnd
            Return sResult
        Catch ex As Exception
            LibError.ErrorView(ex)
            Return retError(sHTMLStart, ex, sHTMLEnd)
        End Try
    End Function

    Public Function SoresenSim(ByVal ODGV As DataGridView, ByVal iColumn1 As Integer, ByVal iColumn2 As Integer) As String

        Dim lRow As Integer = ODGV.RowCount
        Dim iR As Integer
        Dim nSp1 As Integer, nSp2 As Integer, dn1n2 As Double
        Dim sRTF As String = ""
        Dim dSimilar As Double
        Try
            If Auditoria(ODGV) = False Then
                If MsgAuditoria() = False Then
                    Return sRetAuditoria
                    Exit Function
                End If
            End If
            For iR = 0 To lRow - 1
                If ODGV.Rows(iR).Cells(iColumn1).Value > 0 Then
                    nSp1 += 1
                End If
                If ODGV.Rows(iR).Cells(iColumn2).Value > 0 Then
                    nSp2 += 1
                End If
                If ODGV.Rows(iR).Cells(iColumn1).Value > 0 And ODGV.Rows(iR).Cells(iColumn2).Value > 0 Then
                    dn1n2 += 1
                End If
            Next
            dSimilar = (2 * dn1n2) / (nSp1 + nSp2)
            sRTF &= "<ul class='list-group'>"
            sRTF &= "<li class='list-group-item'>Número de indivíduos da amostra " & ODGV.Columns(iColumn1).HeaderCell.Value & ": " & nSp1 & "</li>"
            sRTF &= "<li class='list-group-item'>Número de indivíduos da amostra " & ODGV.Columns(iColumn2).HeaderCell.Value & ": " & nSp2 & "</li>"
            sRTF &= "<li class='list-group-item'>Similaridade de Sorensen (&beta;): " & Math.Round(dSimilar, 4) & "</li>"
            sRTF &= "</ul>"
            sResult = sHTMLStart & "<h3 class='alert alert-success'>Índice de Similaridade Sorensen</h3>" & sRTF & sFooter & sHTMLEnd
            Return sResult
        Catch ex As Exception
            LibError.ErrorView(ex)
            Return retError(sHTMLStart, ex, sHTMLEnd)
        End Try
    End Function

    Public Function Canberra(ByVal ODGV As DataGridView, ByVal iColumn1 As Integer, ByVal iColumn2 As Integer) As String
        Dim lRow As Integer = ODGV.RowCount
        Dim iR As Integer
        Dim nSp1 As Integer, nSp2 As Integer, dn1n2 As Double
        Dim sRTF As String = ""
        Dim dSimilar As Double
        Dim dCel1 As Double, dCel2 As Double
        Try
            If Auditoria(ODGV) = False Then
                If MsgAuditoria() = False Then
                    Return sRetAuditoria
                    Exit Function
                End If
            End If
            For iR = 0 To lRow - 1
                If ODGV.Rows(iR).Cells(iColumn1).Value > 0 Then
                    nSp1 += 1
                End If
                If ODGV.Rows(iR).Cells(iColumn2).Value > 0 Then
                    nSp2 += 1
                End If
                If ODGV.Rows(iR).Cells(iColumn1).Value > 0 And ODGV.Rows(iR).Cells(iColumn2).Value > 0 Then
                    dCel1 = ODGV.Rows(iR).Cells(iColumn1).Value
                    dCel2 = ODGV.Rows(iR).Cells(iColumn2).Value
                    dn1n2 += (Math.Abs(dCel1 - dCel2) / (dCel1 + dCel2))
                End If
            Next
            dSimilar = dn1n2 / (nSp1 + nSp2)
            sRTF &= "<ul class='list-group'>"
            sRTF &= "<li class='list-group-item'>Número de indivíduos da amostra " & ODGV.Columns(iColumn1).HeaderCell.Value & ": " & nSp1 & "</li>"
            sRTF &= "<li class='list-group-item'>Número de indivíduos da amostra " & ODGV.Columns(iColumn2).HeaderCell.Value & ": " & nSp2 & "</li>"
            sRTF &= "<li class='list-group-item'>Similaridade de Canberra: " & Math.Round(dSimilar, 4) & "</li>"
            sRTF &= "</ul>"
            sResult = sHTMLStart & "<h3 class='alert alert-success'>Índice de Similaridade Canberra</h3>" & sRTF & sFooter & sHTMLEnd
            Return sResult
        Catch ex As Exception
            LibError.ErrorView(ex)
            Return retError(sHTMLStart, ex, sHTMLEnd)
        End Try
    End Function
End Class
