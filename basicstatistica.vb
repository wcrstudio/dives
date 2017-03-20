 Public Function pStatBas(ByRef oDGV As DataGridView) As String
        Dim lCol As Long, lRow As Long
        Dim lC As Integer
        Dim lRc As Long
        Dim iRT As Long
        Dim X2 As Double
        Dim n As Long
        Dim EX_EX2 As Double
        Dim EX_EX2Row As Double
        Dim SDRow As Double
        Dim s As Double
        Dim MeGe As Double
        Dim sRTF As String = ""
        Dim dIC As Double
        n = (oDGV.RowCount) * (oDGV.ColumnCount)
        lCol = oDGV.ColumnCount
        lRow = oDGV.RowCount

        Try
            If Auditoria(oDGV) = False Then
                If MsgAuditoria() = False Then
                    Return sRetAuditoria
                    Exit Function
                End If
            End If
            sRTF &= "<div class='row'><div class='col-xs-4 col-lg-4'>Número de espécies: " & CStr(lRow) & "</div>"
            sRTF &= "<div class='col-xs-4 col-lg-4'>Número de amostras: " & CStr(lCol) & "</div>"
            sRTF &= "<div class='col-xs-4 col-lg-4'>Matriz de dados: " & lRow & " x " & lCol & " (" & n & ")" & "</div></div>"
            sRTF &= "<div class='row'><h4 class='alert alert-info'>Média das Espécies</h4>"
            For lRc = 0 To lRow - 1
                For lC = 0 To lCol - 1
                    If oDGV.Rows(lRc).Cells(lC).Value <> String.Empty Then
                        iRT += CDbl(oDGV.Rows(lRc).Cells(lC).Value)
                        X2 += CDbl(oDGV.Rows(lRc).Cells(lC).Value) ^ 2
                    End If
                Next lC
                EX_EX2Row = X2 - ((iRT ^ 2) / lCol)
                SDRow = Math.Sqrt(EX_EX2Row / (lCol - 1))
                dIC = 1.96 * (SDRow / Math.Sqrt(iRT))
                sRTF &= "<ul class='list-group'>"
                sRTF &= "<li class='list-group-item'>Espécie: " & lRc + 1 & " (<em>" & oDGV.Rows(lRc).HeaderCell.Value & "</em>)</li>"
                sRTF &= "<li class='list-group-item'>Média: x = " & Format(iRT / lCol, "###,###,##0.0#") & "</li>"
                sRTF &= "<li class='list-group-item'>Desvio padrão da Espécie: s = " & Format(SDRow, "###,###,##0.0###") & "</li>"
                sRTF &= "<li class='list-group-item'>Variância da Espécie: s² = " & Format((SDRow ^ 2), "###,###,##0.0###") & "</li>"
                sRTF &= "<li class='list-group-item'>Intervalo de confiança: IC = " & Format(iRT / lCol, "###,###,##0.0###") & " &plusmn; " & Format(dIC, "###,###,##0.0####") & "</li>"
                sRTF &= "<li class='list-group-item'>Soma: " & Format(iRT, "###,###,##0.0#") & "</li>"
                sRTF &= "</ul><hr/>"
                iRT = 0

            Next lRc
            sRTF &= "</div>"
            iRT = 0
            X2 = 0
            For lC = 0 To lCol - 1
                For lRc = 0 To lRow - 1
                    If oDGV.Rows(lRc).Cells(lC).Value <> String.Empty Then
                        iRT += CDbl(oDGV.Rows(lRc).Cells(lC).Value)
                      
                        X2 += CDbl(oDGV.Rows(lRc).Cells(lC).Value) ^ 2
                    End If
                Next lRc

            Next lC
            EX_EX2 = X2 - ((iRT ^ 2) / n)
            MeGe = iRT / n
            Dim CV As Double

            s = Math.Sqrt(EX_EX2 / (n - 1))
            dIC = 1.96 * (s / Math.Sqrt(iRT))
            CV = (s / MeGe) * 100
            sRTF &= "<div class='row'><ul class='list-group'><h4 class='alert alert-info'>Média de todas as espécies</h4>"
            sRTF &= "<li class='list-group-item'>Média Geral: &mu; = " & Format(MeGe, "###,###,##0.0###") & "</li>"
            sRTF &= "<li class='list-group-item'>Somatório : x = " & Format(iRT, "###,###,##0.0###") & "</li>"
            sRTF &= "<li class='list-group-item'>Somatório x² = " & Format(X2, "###,###,##0.0###") & "</li>"
            sRTF &= "<li class='list-group-item'>Desvio Padrão: &sigma; = " & Format(Math.Round(s, 4), "###,###,##0.0###") & "</li>"
            sRTF &= "<li class='list-group-item'>Variância: &sigma;² = " & Format(Math.Round(s ^ 2, 4), "###,###,##0.0###") & "</li>"
            sRTF &= "<li class='list-group-item'>Intervalo de Confiança: IC = " & Format(MeGe, "###,###,##0.0###") & " &plusmn; " & Format(dIC, "###,###,##0.0###") & "</li>"
            sRTF &= "<li class='list-group-item'>Coeficiente de Variação: CV = " & Format(CV, "###,###,##0.0###") & "</li>"
            sRTF &= "<li class='list-group-item'>Erro Padrão da Média: s(x) = " & Format(s / Math.Sqrt(n), "###,###,##0.0###") & "</li>"
            sRTF &= "</ul></div>"
            sResult = sHTMLStart & "<h3 class='alert alert-success'>Estatística Básica dos Dados</h3>" & sRTF & sFooter & sHTMLEnd
            Return sResult
        Catch ex As Exception
            MessageBox.Show("Erro: " & ex.Message & vbNewLine & vbNewLine & "Local de Origem: " & ex.StackTrace, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Return retError(sHTMLStart, ex, sHTMLEnd)
        End Try

    End Function
