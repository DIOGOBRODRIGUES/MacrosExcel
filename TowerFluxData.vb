Sub MediaDiaNoite()
colLege = 5
linLege = 6
ColunaVariavel = 5 'alterar para nova variavel
Aba = "N_AVG_day" 'ALTERAR PARA NOVA ANALISE
colunaNova = 3 'ALTERAR PARA NOVA ANALISE
    Do While Not IsEmpty(Cells(linLege, colLege))
      linha = 6
      MediaTemp = 0
      divisor = 0
      linhanova = 2
      linhaNUll = 0
    Do While Not IsEmpty(Range("A" & linha))
      
       If IsEmpty(Range("A" & linha)) And IsEmpty(Range("A" & linha + 1)) Then
            Exit Do
            
        Else
            If Cells(linha, 2) = Cells(linha + 1, 2) Then
                
                If Cells(linha, ColunaVariavel) = "Null" Then
                  
                Else
                    MediaTemp = MediaTemp + Cells(linha, ColunaVariavel)
                    divisor = divisor + 1
                End If
                
            Else
                If MediaTemp = 0 Then
                      ThisWorkbook.Sheets(Aba).Range("A" & linhanova).Value = Cells(linha, 1)
                      ThisWorkbook.Sheets(Aba).Range("B" & linhanova).Value = Cells(linha, 2)
                      ThisWorkbook.Sheets(Aba).Cells(linhanova, colunaNova) = "Null" 'Alterar para nova analise
                      divisor = 0
                      MediaTemp = 0
                      linhanova = linhanova + 1
                    
                Else
                    If Cells(linha, ColunaVariavel) = "Null" Then
                       Call preecherTab(linhanova, linha, MediaTemp, divisor, Aba, colunaNova)
                       
                    Else
                        MediaTemp = MediaTemp + Cells(linha, ColunaVariavel)
                        divisor = divisor + 1
                        Call preecherTab(linhanova, linha, MediaTemp, divisor, Aba, colunaNova)
                       
                    End If
                                    
                End If
                
            End If
            
            linha = linha + 1
            
        End If

    Loop
    colunaNova = colunaNova + 1
    ColunaVariavel = ColunaVariavel + 1
    colLege = colLege + 1
Loop
End Sub

Sub preecherTab(linhanova, linha, MediaTemp, divisor, Aba, colunaNova)
    MediaDiaria = MediaTemp / divisor
    ThisWorkbook.Sheets(Aba).Range("A" & linhanova).Value = Cells(linha, 1)
    ThisWorkbook.Sheets(Aba).Range("B" & linhanova).Value = Cells(linha, 2)
    ThisWorkbook.Sheets(Aba).Cells(linhanova, colunaNova) = MediaDiaria 'Alterar para nova analise
    divisor = 0
    MediaTemp = 0
    linhanova = linhanova + 1
End Sub



