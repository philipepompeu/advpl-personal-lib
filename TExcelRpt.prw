#INCLUDE 'TOTVS.CH'

/*/{Protheus.doc} TExcelRpt
	Classe criada para facilitar a criacao de novos
	relatorios simples
@author PHILIPE.POMPEU
@since 05/09/2017
@version 1.0
@example
(examples)
@see (links_or_references)
/*/CLASS TExcelRpt
	Data cTitulo as String
	Data cPergunte as String
	METHOD New(cPerg,cTitle) CONSTRUCTOR	
	/*Ao chamar o metodo Execute ele obtera query de GetQuery e automaticamente
	exportara pro Excel*/
	METHOD Execute()
	METHOD ToExcel(cQuery)
	METHOD ExcelQuery(cUmAlias,cQuery)
	METHOD GetDeAte(dDtDe, dDtAte, cCampo)	
	/*Os metodos FixSX1 e GetQuery devem ser sobrescritos pelas classes filhas*/
	METHOD FixSX1()
	METHOD GetQuery()
ENDCLASS

//-----------------------------------------------------------------
METHOD New(cPerg,cTitle) CLASS TExcelRpt
	Default cTitle := "Exportação p/ Excel"
	::cPergunte := cPerg
	::cTitulo := cTitle
Return

Method FixSX1() CLASS TExcelRpt
Return

Method GetQuery() CLASS TExcelRpt
	Local cQuery := ""
Return cQuery

Method Execute() CLASS TExcelRpt
	Local aArea 	:= GetArea()
	Local bExp := {|x|::ToExcel()}
	::FixSX1()
	if(Pergunte(::cPergunte))
		MakeSqlExpr(::cPergunte)
		
		if(IsBlind())
			Eval(bExp)
		else
			FWMsgRun( , bExp, "Gerando relatório", "Processando..." )
		endIf
	endIf
	RestArea(aArea)
Return

/*/{Protheus.doc} ToExcel
	Rotina que exporta para o Excel uma consulta qualquer
@author PHILIPE.POMPEU
@since 07/08/2017
@return ${return}, ${return_description}
/*/Method ToExcel(cQuery) CLASS TExcelRpt
	Local cMyAlias	:= GetNextAlias()
	Local aResult := {}
	Default cQuery := ::GetQuery()
	aResult := ::ExcelQuery(cMyAlias, cQuery)	
	
	DlgToExcel({{"GETDADOS",::cTitulo,aResult[1],aResult[2]}})
	
	aSize(aResult[2],0)
	aSize(aResult,0)
Return

Method ExcelQuery(cUmAlias,cQuery) CLASS TExcelRpt
	Local aCpo := {}
	Local aTemp := {}
	Local aResult := {}
	Local aCab := {}
	Local aDados:= {}
	Local aTemp := {}
	Local nI := 0
	Local cTitulo := ""
	Local cPict :=""
	Local xValor
	Local aStruct	:= {}
	Local nPos := 0
	Local cTipo := ''
	Local nLen := 0
	Local aSm0	:= {}
	
	cQuery	:= ChangeQuery(cQuery)				
	DbUseArea(.T., "TOPCONN", TCGenQry(,,cQuery), cUmAlias, .T., .F.)
	
	aStruct := (cUmAlias)->(DbStruct())
	
	SX3->(DbSetOrder(2))//X3_CAMPO
	
	for nI:= 1 to Len(aStruct)
		if(SX3->(DbSeek(aStruct[nI,1])))
			if(AllTrim(SX3->X3_TIPO) $ 'D|N')
				TCSetField(cUmAlias, aStruct[nI,1], AllTrim(SX3->X3_TIPO), SX3->X3_TAMANHO, SX3->X3_DECIMAL)	
			endIf
			cPict := AllTrim(SX3->X3_CBOX)
			if(Empty(cPict))
				if('_FILIAL' $ AllTrim(aStruct[nI,1]))
					if(Len(aSm0) <= 0)
						cPict := {}
						aSm0 := FWLoadSm0()
						for nJ:= 1 to Len(aSm0)
							aAdd(cPict,{AllTrim(aSm0[nJ,7]),aSm0[nJ,2],AllTrim(aSm0[nJ,2]) + ' - ' + AllTrim(aSm0[nJ,7])})
						next nJ
						aSm0 := aClone(cPict)
					endIf
					cPict := aSm0
				else				
					cPict := SX3->X3_PICTURE
				endIf
			else
				cPict := RetSx3Box(X3CBox(),,, 1)	
			endIf
			aAdd(aCab,{aStruct[nI,1],"C", SX3->X3_TAMANHO,SX3->X3_DECIMAL,SX3->X3_TITULO,cPict})
		else
			aAdd(aCab,{aStruct[nI,1],"C", aStruct[nI,3],aStruct[nI,4],aStruct[nI,1],""})
		endIf		
	next nI
	
	nLen := Len(aCab)
	
	while ( (cUmAlias)->(!Eof()) )
		aTemp := {}	
		for nI:= 1 to nLen
			xValor := (cUmAlias)->&(aCab[nI,1])
			cTipo := ValType(xValor)
			if(cTipo == 'D')
				xValor := DToC(xValor)
			elseIf(ValType(aCab[nI,6]) == 'C')
				cPict := AllTrim(aCab[nI,6])
				
				if!(Empty(cPict))					
					xValor := Transform(xValor,cPict)
				elseIf(cTipo == 'N')
					xValor := cValToChar(xValor)
				endIf
				
			elseIf(ValType(aCab[nI,6]) == 'A') //Nesse caso trata-se de um campo ComboBox
				If(cTipo == 'N')
					xValor := cValToChar(xValor)				
				endIf
				xValor := AllTrim(xValor)
				nPos := aScan(aCab[nI,6],{|x|AllTrim(x[2]) == xValor})
				if(nPos > 0)
					xValor := AllTrim(aCab[nI,6,nPos,3])
				endIf
			endIf
			
			if(cTipo == 'C')
				xValor := CHR(160) + xValor
			endIf
			
			aAdd(aTemp,xValor)
		next nI
		aSize(aTemp,Len(aTemp)+1)
		aAdd(aDados,aClone(aTemp))
	
		(cUmAlias)->(dbSkip())
	EndDo		
	(cUmAlias)->(dbCloseArea())
	
	for nI:= 1 to nLen
		cTitulo := AllTrim(aCab[nI,5])		
		aCab[nI,1] := Upper(cTitulo)
		
		aSize(aCab[nI],4) //Remove os dois últimos campos que são usados apenas pra essa rotina
	next nI
	aSize(aCab,Len(aCab)+1)
	aTail(aCab) := {" " ," ", 0, 0}
	
	aAdd(aResult,aCab)
	aAdd(aResult,aDados)
				 
Return aResult

METHOD GetDeAte(dDtDe, dDtAte, cCampo) CLASS TExcelRpt
	Local cDataDeAte := '' 
	
	if!(Empty(dDtDe) .Or. Empty(dDtAte))		
		cDataDeAte:= "("+ cCampo +" BETWEEN '"+ DtoS(dDtDe) +"' AND '"+ DtoS(dDtAte) +"')"
	else
		if(!Empty(dDtDe) .and. Empty(dDtAte))
			cDataDeAte:= "("+ cCampo +" >= '"+ DtoS(dDtDe) +"')"
		elseIf(!Empty(dDtAte) .and. Empty(dDtDe))
			cDataDeAte:= "("+ cCampo +" <= '"+ DtoS(dDtAte) +"')"
		else
			cDataDeAte:= "("+ cCampo +" >= '20000101')"
		endIf
	endIf
Return cDataDeAte