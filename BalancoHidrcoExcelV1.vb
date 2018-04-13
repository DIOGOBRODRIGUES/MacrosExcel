Sub balancohidrico()
    colunapetp = 26
    colunanac = 38
    colunaarm = 50
	
    '1-definindo inicio do balanco
	somaPEtp = 0
	somaPEtpPos = 0
	somaPEtpNeg = 0
	inicio = 1
	
	For cetp = colunapetp to 37
		somaPEtp = somaPEtp +Cells(4, cetp)
		If Cells(4, cetp) > 0 Then
			somaPEtpPos = somaPEtpPos + Cells (4, cetp)
		Else
			somaPEtpNeg = somaPEtpNeg + Cells (4, cetp)
		End If
		
		If Cells(4, cetp) > 0 and Cells(4, cetp+1)< 0 Then
				inicio = Cells(3, cetp)
		
		End If
		

	Next cetp
		
		If somaPEtp > 0 Then
			Cells(4, colunanac + inicio - 1) = 0
			Cells (4, colunaarm+inicio - 1) = Cells(4, 1)
		Else
			If somaPEtpPos > Cells(4, 1) Then
				Cells(4, colunanac + inicio - 1) = 0
				Cells (4, colunaarm+inicio - 1)= Cells(4, 1)
			Else
				Cells(4, colunanac + inicio - 1) = Cells(4, 1)*Application.WorksheetFunction.Ln((somaPEtpPos/Cells(4, 1))/(1-Exp(somaPEtpNeg/Cells(4,1))))
				Cells(4, colunaarm + inicio - 1) = Cells(4, 1)*Exp(-Abs(Cells(4, colunanac + inicio - 1)/Cells(4, 1)))
			End If
		End If
			

    '******************************Fim do Algortimo de inicialização
	'2-Preenchendo o NAC e ARM
	colunanac = colunanac + inicio 
    colunaarm = colunaarm +inicio 
	colunapetp  =colunapetp + inicio 

	
	For coluna = colunapetp To 37
		
	If Cells (4, colunaarm) <> Empty Then
            Exit For
	End If
		
		If Cells (4, colunapetp) < 0 Then
			Cells (4, colunanac) = Cells(4, colunapetp) +  Cells (4, colunanac-1)
			Cells (4 , colunaarm) = Cells(4, 1)*Exp(-Abs(Cells (4, colunanac)/Cells(4, 1)))
		
		Else
			Cells (4, colunaarm) = Cells(4, colunapetp) +  Cells (4, colunaarm-1)
			If Cells (4, colunaarm) > Cells(4, 1) Then
				Cells (4, colunaarm) = Cells(4, 1)
			End If
			Cells (4, colunanac) = Cells(4, 1) *Application.WorksheetFunction.Ln(Cells (4, colunaarm)/Cells(4, 1))
		End If
		
		'loop de Dezembro para Janeiro
		If Cells (3 , colunaarm) = 12 Then
			  colunapetp = 26
			  colunanac = 38
			  colunaarm = 50
			  coluna = colunapetp
			If Cells (4, colunapetp) < 0 Then
				Cells (4, colunanac) = Cells(4, colunapetp) +  Cells (4, colunanac+11)
				Cells (4 , colunaarm) = Cells(4, 1)*Exp(-Abs(Cells (4, colunanac)/Cells(4, 1)))
		
			Else
				Cells (4, colunaarm) = Cells(4, colunapetp) +  Cells (4, colunaarm+11)
				If Cells (4, colunaarm) > Cells(4, 1) Then
					Cells (4, colunaarm) = Cells(4, 1)
				End If
				Cells (4, colunanac) = Cells(4, 1) *Application.WorksheetFunction.Ln(Cells (4, colunaarm)/Cells(4, 1))
			End If
		End If
		colunanac = colunanac + 1
        colunaarm = colunaarm + 1
		colunapetp = colunapetp + 1
		
	
	Next coluna
	
	
	
	'3-Calculo da Alteração Alt
	colunapetp = 26
    colunanac = 38
    colunaarm = 50
	colunaAlt = 63
	mes = 1
	  'preenchendo janeiro
	Cells(4, 62) = Cells(4, colunaarm) - Cells(4 , colunaarm + 11)
		'Resto dos meses
	For cAlt = colunaAlt to 73
			Cells (4 ,cAlt) = Cells(4, colunaarm + mes) - Cells( 4, colunaarm + mes -1)
			mes = mes +1
       
    Next cAlt 
	
	
	
	'4-Cálculo da ETR
	mes = 0
	colunaP =  14
	colunaEtp = 2
	colunaAlt = 62
	colunaEtr = 74
	
	For cEtr = colunaEtr to 85
		If Cells( 4, colunapetp + mes) < 0 Then
			Cells (4, cEtr) = Cells ( 4, colunaP + mes)+ Abs(Cells(4, colunaAlt + mes))
		Else
			Cells(4, cEtr) = Cells(4, colunaEtp + mes)
		End If
		
		mes = mes + 1
		
	Next cEtr
	
	'5 - Determinar DEF
	
	colunaDef = 86
	mes = 0
	For cDef = colunaDef to 97
		Cells(4, cDef) = Cells(4, colunaEtp + mes ) - Cells(4, colunaEtr + mes )
		mes = mes + 1
	
	Next cDef
	
	' 6 Determinar o Exc
	
	colunaExc = 98
	mes = 0
	For cExc = colunaExc to 109
		If Cells(4, colunaarm + mes) < Cells (4, 1)  Then
			Cells(4, cExc) = 0
		End If
			
		If	Cells(4, colunaarm + mes) = Cells(4, 1) Then	
			Cells (4, cExc) = Cells(4, colunapetp + mes )- Cells(4, colunaAlt + mes)
		End If
		mes = mes + 1
	Next cExc
	
	
End Sub