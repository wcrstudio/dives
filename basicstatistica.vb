Imports System.Windows.Forms

Public Class basicstatistica
    Public Sub New()
        Main()
    End Sub
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
Public Function PearsonCorrel(ByRef oDGV As DataGridView, ByVal iColumn1 As Integer, ByVal iColumn2 As Integer) As String
        Dim sRTF As String = ""
        Dim dPearson As Double
        Dim lCol As Long, lRow As Long
        Dim Exy As Double, Ex As Double, Ey As Double, Ex2 As Double, Ey2 As Double
        Dim nSp1 As Integer, nSp2 As Integer, nSpCom As Integer
        Dim TCorrel As Double, GL As Double, lAlpha As Double = 5
        Try
            lCol = oDGV.ColumnCount
            lRow = oDGV.RowCount

            If Auditoria(oDGV) = False Then
                If MsgAuditoria() = False Then
                    Return sRetAuditoria
                    Exit Function
                End If
            End If
            For iR = 0 To lRow - 1
                If oDGV.Rows(iR).Cells(iColumn1).Value > 0 Then
                    nSp1 += oDGV.Rows(iR).Cells(iColumn1).Value
                   
                End If
                If oDGV.Rows(iR).Cells(iColumn2).Value > 0 Then
                    nSp2 += oDGV.Rows(iR).Cells(iColumn2).Value
                    
                End If
              
                'End If
                Ex += oDGV.Rows(iR).Cells(iColumn1).Value
                Ex2 += (oDGV.Rows(iR).Cells(iColumn1).Value) ^ 2
                Ey += oDGV.Rows(iR).Cells(iColumn2).Value
                Ey2 += (oDGV.Rows(iR).Cells(iColumn2).Value) ^ 2
                Exy += (oDGV.Rows(iR).Cells(iColumn1).Value * oDGV.Rows(iR).Cells(iColumn2).Value)
            Next
            nSpCom = nSp1 + nSp2
            GL = lRow - 2
            dPearson = (nSpCom * (Exy) - (Ex * Ey)) / (((nSpCom * (Ex2) - (Ex) ^ 2) ^ (1 / 2)) * (nSpCom * (Ey2) - (Ey) ^ 2) ^ (1 / 2))
            TCorrel = Math.Abs(dPearson * (Math.Sqrt(GL / (1 - dPearson ^ 2))))
            sRTF &= "<ul class='list-group'>"
            sRTF &= "<li class='list-group-item'>Número de indivíduos da amostra " & oDGV.Columns(iColumn1).HeaderCell.Value & ": " & nSp1 & "</li>"
            sRTF &= "<li class='list-group-item'>Número de indivíduos da amostra " & oDGV.Columns(iColumn2).HeaderCell.Value & ": " & nSp2 & "</li>"
            sRTF &= "<li class='list-group-item'>Número total de indivíduos: " & nSpCom & "</li>"
            sRTF &= "<li class='list-group-item'>Correlação de Pearson: " & Math.Round(dPearson, 4) & "</li>"
            If dPearson < 0 Then
                sRTF &= "<li class='list-group-item list-group-success'>Correlação negativa</li>"
            ElseIf dPearson = 0 Then
                sRTF &= "<li class='list-group-item list-group-item-warning'>Correlação neutra</li>"
            ElseIf dPearson > 1 Then
                sRTF &= "<li class='list-group-item list-group-success'>Correlação positiva</li>"
            End If
            sRTF &= "<li class='list-group-item'>Grau de liberdade (n-2): " & GL & "</li>"
            sRTF &= "<li class='list-group-item'>Teste t-<em>student</em> para r: " & Math.Round(TCorrel, 4) & "</li>"
            Dim tTabV As Double = TTab(lAlpha, Math.Round(GL, 0))
            sRTF &= "<li class='list-group-item'>Valor T Tabelado t(&alpha;=0,05)(2, " & Math.Round(GL, 0) & "): " & tTabV & "</li>"

            If Math.Round(TCorrel, 4) >= tTabV Then
                sRTF &= "<li class='list-group-item list-group-item-success'><strong>HÁ</strong> correlação significativa entre os pares de dados, segundo o teste  t-<em>student</em> para r a 5% de probabilidade</li>"
            Else
                sRTF &= "<li class='list-group-item list-group-item-success'><strong>NÃO HÁ</strong> correlação significativa entre os pares de dados, segundo o teste  t-<em>student</em> para r a 5% de probabilidade</li>"
            End If
            sRTF &= "</ul>"


            sResult = sHTMLStart & "<h3 class='alert alert-success'>Correlação de Pearson</h3>" & sRTF & sFooter & sHTMLEnd
            Return sResult
        Catch excpt As Exception
            MessageBox.Show("Erro: " & excpt.Message & vbNewLine & vbNewLine & "Local de Origem: " & excpt.StackTrace, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Return retError(sHTMLStart, excpt, sHTMLEnd)
        End Try
    End Function
  
  Protected Overrides Sub Finalize()
        MyBase.Finalize()
  End Sub
End Class
