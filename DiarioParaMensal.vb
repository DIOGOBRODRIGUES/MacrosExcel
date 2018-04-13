Sub Botão1_Clique()
    linha = 2
    linhaMes = 2
    somaMesX = 0
    somaMesY = 0
    
    Do While Not IsEmpty(Range("A" & linha))
    
        If Cells(linha + 1, 1) <> "" Then
        
            Do While Month(Cells(linha, 1)) = Month(Cells(linha + 1, 1))
                somaMesX = somaMesX + Cells(linha, 2)
                somaMesY = somaMesY + Cells(linha, 3)
                linha = linha + 1
            Loop
            somaMesX = somaMesX + Cells(linha, 2)
            somaMesY = somaMesY + Cells(linha, 3)
            Cells(linhaMes, 8) = Cells(linha, 1)
            Cells(linhaMes, 8).NumberFormat = "mm/yyyy"
            Cells(linhaMes, 9) = somaMesX
            Cells(linhaMes, 10) = somaMesY
            somaMesX = 0
            somaMesY = 0
            linha = linha + 1
            linhaMes = linhaMes + 1
        End If
        
    Loop
