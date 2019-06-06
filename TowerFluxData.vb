Sub MediaDiaNoite()
    linha = 2 'ALTERAR PARA NOVA ANALISE
    MediaTemp = 0
    divisor = 0
    linhanova = 2
    linhaNUll = 0
    colunaDaVariavel = 25 'ALTERAR PARA NOVA ANALISE
    ColunaAno = 1 'ALTERAR PARA NOVA ANALISE
    ColunaDia = 2 'ALTERAR PARA NOVA ANALISE
    Aba = "N_AVG_day" 'ALTERAR PARA NOVA ANALISE
    
	
    Do While Not IsEmpty(Range("A" & linha))
      
       If IsEmpty(Range("A" & linha)) And IsEmpty(Range("A" & linha + 1)) Then
            Exit Do
            
        Else
            If Cells(linha, ColunaDia) = Cells(linha + 1, ColunaDia) Then
                
                If Cells(linha, colunaDaVariavel) = "Null" Or Cells(linha, colunaDaVariavel) < 0 Then
                  
                Else
                    MediaTemp = MediaTemp + Cells(linha, colunaDaVariavel)
                    divisor = divisor + 1
                End If
                
            Else
                If MediaTemp <= 0 Then 'ALTERADO NA CHUVA
                      'ThisWorkbook.Sheets(Aba).Range("A" & linhanova).Value = Cells(linha, ColunaAno)
                      'ThisWorkbook.Sheets(Aba).Range("B" & linhanova).Value = Cells(linha, ColunaDia)
                      ThisWorkbook.Sheets(Aba).Range("X" & linhanova).Value = "Null"      'ALTERAR PARA NOVA ANALISE
                      divisor = 0
                      MediaTemp = 0
                      linhanova = linhanova + 1
                    
                Else
                    
                    If Cells(linha, colunaDaVariavel) = "Null" Or Cells(linha, colunaDaVariavel) <= 0 Then 'ALTERADO NA CHUVA
                       Call preecherTab(linhanova, linha, MediaTemp, divisor, ColunaAno, ColunaDia, Aba)
                       
                    Else
                        
                        MediaTemp = MediaTemp + Cells(linha, colunaDaVariavel)
                        divisor = divisor + 1
                        Call preecherTab(linhanova, linha, MediaTemp, divisor, ColunaAno, ColunaDia, Aba)
                       
                    End If
                                    
                End If
                
            End If
            
            linha = linha + 1
            
        End If

    Loop

End Sub

Sub preecherTab(linhanova, linha, MediaTemp, divisor, ColunaAno, ColunaDia, Aba)
    MediaDiaria = MediaTemp / divisor
    'ThisWorkbook.Sheets(Aba).Range("A" & linhanova).Value = Cells(linha, ColunaAno)
    'ThisWorkbook.Sheets(Aba).Range("B" & linhanova).Value = Cells(linha, ColunaDia)
    ThisWorkbook.Sheets(Aba).Range("X" & linhanova).Value = MediaDiaria     'ALTERAR PARA NOVA ANALISE
    divisor = 0
    MediaTemp = 0
    linhanova = linhanova + 1
End Sub