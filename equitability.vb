Imports System.Windows.Forms

Public Class equitability
    Dim dSimp As New diversity
    Dim dSW As New diversity
    Public ValueLog As Double
    Public Sub New()
        Main()
    End Sub

    Public Function EqJ(ByRef oDGV As DataGridView) As String
        Dim lcol As Long, lrow As Long
        Dim iR As Long, iC As Integer
        Dim lCT As Long, lCel As Long
        Dim iCelValue As Integer
        Dim lSP As Double, lEJ As Double
        Dim sLev As String, lHMax As Double, lSW As Double
        Dim sRTF As String = ""

        lcol = oDGV.ColumnCount
        lrow = oDGV.RowCount
        Try
            If Auditoria(oDGV) = False Then
                If MsgAuditoria() = False Then
                    Return sRetAuditoria
                    Exit Function
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
                sRTF &= "<div class='table-responsive'><table class='table table-bordered table-hover'><thead><tr class='info'><th>Espécie</th><th>Número de Indivíduos</th></tr></thead><tbody>"

                For iCelValue = 0 To lrow - 1
                    If oDGV.Rows(iCelValue).Cells(iC).Value <> String.Empty Then
                        lCel = CDbl(oDGV.Rows(iCelValue).Cells(iC).Value)
                    Else
                        lCel = 0
                    End If
                    If lCel > 0 Then
                        sRTF &= "<tr>"
                        sRTF &= "<td>" & oDGV.Rows(iCelValue).HeaderCell.Value & "</td>"
                        sRTF &= "<td class='text-center'>" & lCel & "</td></tr>"
                        If lCel <> 0 Then
                            lSP += 1
                        End If
                    End If
                Next iCelValue
                sRTF &= "</tbody><tfooter><tr class='text-info'>"
                sRTF &= "<td>Número de espécies: " & lSP & "</td>"
                lHMax = Math.Log10(lSP)
                sRTF &= "<td>Diversidade Máxima para o Índice de Shannon (H<sub>max</sub>): " & Math.Round(lHMax, 4) & "</td></tr>"
                lSW = dSW.fDivSW(oDGV, iC)
                sRTF &= "<tr class='text-info'><td>Índice Diversidade de Shannon-Wiener: " & Math.Round(lSW, 4) & "</td>"

                lEJ = lSW / lHMax
                sRTF &= "<td>Índice Equidade de J (Shannon-Wiener): " & Format(Math.Abs(lEJ), "###,###,###,##0.0###") & "</td>"
                sRTF &= "</tr></tfooter></table></div><hr/>"
                lHMax = 0
                lSP = 0
            Next iC
            sResult = sHTMLStart & "<h3 class='alert alert-success'>Relatório da Equidade J (Shannon-Wiener)</h3>" & sRTF & sFooter & sHTMLEnd
            Return sResult
        Catch ex As Exception
            LibError.ErrorView(ex)
            Return retError(sHTMLStart, ex, sHTMLEnd)
        End Try
    End Function

    Public Function EqD(ByRef oDGV As DataGridView) As String
        Dim lcol As Long, lrow As Long
        Dim iR As Long, iC As Integer
        Dim lCT As Long, lCel As Long
        Dim iCelValue As Integer
        Dim lSP As Double, lED As Double
        Dim sLev As String, lDMax As Double, lSimp As Double
        Dim sRTF As String = ""
        lcol = oDGV.ColumnCount
        lrow = oDGV.RowCount

        Try
            If Auditoria(oDGV) = False Then
                If MsgAuditoria() = False Then
                    Return sRetAuditoria
                    Exit Function
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
                sRTF &= "<div class='table-responsive'><table class='table table-bordered table-hover'><thead><tr class='info'><th>Espécie</th><th>Número de Indivíduos</th></tr></thead><tbody>"
                For iCelValue = 0 To lrow - 1
                    If oDGV.Rows(iCelValue).Cells(iC).Value <> String.Empty Then
                        lCel = CDbl(oDGV.Rows(iCelValue).Cells(iC).Value)
                        sRTF &= "<tr>"
                        sRTF &= "<td>" & oDGV.Rows(iCelValue).HeaderCell.Value & "</td>"
                        sRTF &= "<td class='text-center'>" & lCel & "</td></tr>"
                        If lCel > 0 Then lSP += 1
                    End If
                Next iCelValue
                sRTF &= "</tbody><tfooter><tr class='text-info'>"
                sRTF &= "<td>Número de espécies: " & lSP & "</td>"
                lDMax = ((lSP - 1) / lSP) * (lCT / (lCT - 1))
                sRTF &= "<td>Diversidade Máxima para o Índice de Simpson (D<sub>max</sub>): " & Math.Round(lDMax, 4) & "</td></tr>"
                lSimp = dSimp.fdivSimpson(oDGV, iC)
                sRTF &= "<tr class='text-info'><td>Índice Diversidade de Simpson: " & Math.Round(lSimp, 4) & "</td>"
                lED = lSimp / lDMax
                sRTF &= "<td>Índice Equidade de ED (Simpson): " & Format(Math.Abs(lED), "###,###,###,##0.0###") & "</td>"
                sRTF &= "</tr></tfooter></table></div><hr/>"
                lDMax = 0
                lSP = 0
            Next iC

            sResult = sHTMLStart & "<h3 class='alert alert-success'>Relatório da Equidade ED (Simpson)</h3>" & sRTF & sFooter & sHTMLEnd
            Return sResult
        Catch ex As Exception
            LibError.ErrorView(ex)
            Return retError(sHTMLStart, ex, sHTMLEnd)
        End Try
    End Function

    Public Function EHill(ByRef oDGV As DataGridView) As String
        Dim lcol As Long, lrow As Long
        Dim iR As Long, iC As Integer
        Dim lCT As Long, lCel As Long
        Dim iCelValue As Integer
        Dim lSP As Double, lEHill As Double
        Dim sLev As String, lDs As Double, lSW As Double
        Dim sRTF As String = ""
        lcol = oDGV.ColumnCount
        lrow = oDGV.RowCount
        Try
            If Auditoria(oDGV) = False Then
                If MsgAuditoria() = False Then
                    Return sRetAuditoria
                    Exit Function
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
                sRTF &= "<div class='table-responsive'><table class='table table-bordered table-hover'><thead><tr class='info'><th>Espécie</th><th>Número de Indivíduos</th></tr></thead><tbody>"
                For iCelValue = 0 To lrow - 1
                    If oDGV.Rows(iCelValue).Cells(iC).Value <> String.Empty Then
                        lCel = CDbl(oDGV.Rows(iCelValue).Cells(iC).Value)
                        sRTF &= "<tr>"
                        sRTF &= "<td>" & oDGV.Rows(iCelValue).HeaderCell.Value & "</td>"
                        sRTF &= "<td class='text-center'>" & lCel & "</td></tr>"
                        If lCel <> 0 Then lSP += 1
                    End If

                Next iCelValue
                lDs = dSimp.fdivSimpson(oDGV, iC)
                lSW = dSW.fDivSW(oDGV, iC)
                sRTF &= "</tbody><tfooter><tr class='text-info'>"
                sRTF &= "<td>Índice de Diversidade de Shannon-Wiener (H'): " & Math.Round(lSW, 4) & "</td>"
                sRTF &= "<td>Índice de Diversidade de Simpson (Ds): " & Math.Round(lDs, 4) & "</td></tr>"
                sRTF &= "<tr class='text-info'><td>(1/Ds) - 1 : " & Math.Round((1 / lDs) - 1, 4) & "</td>"
                sRTF &= "<td>e<sup>H'</sup>: " & Math.Round((Math.E ^ lSW), 4) & "</td></tr>"

                lEHill = ((1 / lDs) - 1) / ((Math.E ^ lSW) - 1)

                sRTF &= "<tr class='text-info'><td colspan=2>Índice Equidade de Hill  Modificado (EH): " & Format(Math.Abs(lEHill), "###,###,###,##0.0###") & "</td>"
                sRTF &= "</tr></tfooter></table></div><hr/>"
                lEHill = 0
            Next iC
            sResult = sHTMLStart & "<h3 class='alert alert-success'>Relatório da Equidade de Hill (Modificado)</h3>" & sRTF & sFooter & sHTMLEnd
            Return sResult
        Catch ex As Exception
            LibError.ErrorView(ex)
            Return retError(sHTMLStart, ex, sHTMLEnd)
        End Try
    End Function

    Public Function EqE(ByRef oDGV As DataGridView) As String
        Dim lcol As Long, lrow As Long
        Dim iR As Integer, iC As Integer
        Dim lCT As Long, lCel As Long
        Dim iCelValue As Integer
        Dim sLev As String, lEqE As Double
        Dim sRTF As String = ""
        Dim lSp As Integer, lSW As Double
        lcol = oDGV.ColumnCount
        lrow = oDGV.RowCount
        Try
            If Auditoria(oDGV) = False Then
                If MsgAuditoria() = False Then
                    Return sRetAuditoria
                    Exit Function
                End If
            End If

            For iC = 0 To lcol - 1 Step 1
                lCT = 0
                For iR = 0 To lrow - 1 Step 1
                    'numero total de indivíduos
                    If oDGV.Rows(iR).Cells(iC).Value <> String.Empty Then _
                    lCT += CDbl(oDGV.Rows(iR).Cells(iC).Value)
                Next iR

                sLev = oDGV.Columns(iC).HeaderText.ToString
                sRTF &= "<div class='row'><div class='col-xs-6 col-lg-6'>Levantamento: " & sLev & "</div>"
                sRTF &= "<div class='col-xs-6 col-lg-6'>Valor N: " & lCT & "</div></div>"
                sRTF &= "<div class='table-responsive'><table class='table table-bordered table-hover'><thead><tr class='info'><th>Espécie</th><th>Número de Indivíduos</th></tr></thead><tbody>"

                For iCelValue = 0 To lrow - 1
                    If oDGV.Rows(iCelValue).Cells(iC).Value <> String.Empty Then
                        lCel = CDbl(oDGV.Rows(iCelValue).Cells(iC).Value)
                        If lCel <> 0 Then lSp += 1
                        sRTF &= "<tr>"
                        sRTF &= "<td>" & oDGV.Rows(iCelValue).HeaderCell.Value & "</td>"
                        sRTF &= "<td class='text-center'>" & lCel & "</td></tr>"
                    End If

                Next iCelValue

                lSW = dSW.fDivSW(oDGV, iC)
                lEqE = (Math.E ^ lSW) / lSp
                sRTF &= "</tbody><tfooter><tr class='text-info'>"
                sRTF &= "<td>Número de espécies (S): " & lSp & "</td>"
                sRTF &= "<td>Diversidade de Espécie H' (Shannon-Wiener): " & Math.Round(lSW, 4) & "</td></tr>"
                sRTF &= "<tr class='text-info'><td colspan=2>Valor da Equidade (E): " & Math.Round(lEqE, 4) & "</td></tr>"
                sRTF &= "</table></div><hr/>"
                lSW = 0
                lEqE = 0
                lSp = 0
            Next iC

            sResult = sHTMLStart & "<h3 class='alert alert-success'>Relatório da Equidade E</h3>" & sRTF & sFooter & sHTMLEnd
            Return sResult
        Catch ex As Exception
            LibError.ErrorView(ex)
            Return retError(sHTMLStart, ex, sHTMLEnd)
        End Try
    End Function

    Public Function EqU(ByRef oDGV As DataGridView) As String
        Dim lcol As Long, lrow As Long
        Dim iR As Long, iC As Integer
        Dim lCT As Long, lCel As Long
        Dim lCel2 As Double
        Dim iCelValue As Integer
        Dim EqUValue As Double, sLev As String
        Dim dU As Double, lSp As Integer
        Dim sRTF As String = ""
        lcol = oDGV.ColumnCount
        lrow = oDGV.RowCount
        Try
            If Auditoria(oDGV) = False Then
                If MsgAuditoria() = False Then
                    Return sRetAuditoria
                    Exit Function
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
                sRTF &= "<div class='table-responsive'><table class='table table-bordered table-hover'><thead><tr class='info'><th>Espécie</th><th>Número de Indivíduos</th></tr></thead><tbody>"
                For iCelValue = 0 To lrow - 1
                    If oDGV.Rows(iCelValue).Cells(iC).Value <> String.Empty Then
                        lCel = CDbl(oDGV.Rows(iCelValue).Cells(iC).Value)
                        If lCel <> 0 Then lSp += 1
                        lCel2 += lCel ^ 2
                        sRTF &= "<tr>"
                        sRTF &= "<td>" & oDGV.Rows(iCelValue).HeaderCell.Value & "</td>"
                        sRTF &= "<td class='text-center'>" & lCel & "</td></tr>"
                    End If
                Next iCelValue
                dU = Math.Sqrt(lCel2)
                EqUValue = (lCT - dU) / (lCT - (lCT / (Math.Sqrt(lSp))))
                sRTF &= "</tbody><tfooter><tr class='text-info'>"
                sRTF &= "<td>Número de Espécie: " & lSp & "</td>"
                sRTF &= "<td>Valor de U: " & Math.Round(dU, 4) & "</td></tr>"
                sRTF &= "<tr class='text-info'><td colspan=2>Índice Equidade de McIntosh (EqU): " & Format(Math.Abs(EqUValue), "###,###,###,##0.0###") & "</td></tr>"
                sRTF &= "</table></div><hr/>"


                dU = 0
                lCel2 = 0
                lSp = 0
            Next iC
            sResult = sHTMLStart & "<h3 class='alert alert-success'>Relatório da Equidade de McIntosh</h3>" & sRTF & sFooter & sHTMLEnd
            Return sResult
        Catch ex As Exception
            LibError.ErrorView(ex)
            Return retError(sHTMLStart, ex, sHTMLEnd)
        End Try
    End Function
End Class
