Sub formatarAnual()
    l = 2
    coluna = 6
    linha = 2
    
    Do While Not IsEmpty(Range("A" & l))
        Cells(linha, 5) = Cells(l, 1)
        Do While Cells(l, 1) = Cells(l + 1, 1)
            Cells(linha, coluna) = Cells(l, 3)
            l = l + 1
            coluna = coluna + 1
        Loop
        If Cells(l, 2) = 12 Then
            Cells(linha, coluna) = Cells(l, 3)
        End If
        coluna = 6
        linha = linha + 1
        l = l + 1
        
    Loop
        
End Sub