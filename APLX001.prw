#INCLUDE 'PROTHEUS.CH'

/*Variavel utilizada nas funcoes GetOptX3 e RetOptX3*/
Static cRetGrupo

/*/{Protheus.doc} RetOptX3
	Retorna os valores selecionados na funcao <GetOptX3>
@author PHILIPE.POMPEU
@since 22/08/2017
@return ${return}, ${return_description}
/*/User Function RetOptX3()
	Local cResult := ""
	cResult := cRetGrupo
	cRetGrupo := ""
Return cResult

/*/{Protheus.doc} GetOptX3
	Mostra uma tela para selecionar as opções de campos combobox
@author PHILIPE.POMPEU
@since 22/08/2017
@param cCampo, character, (Descrição do parâmetro)
@return ${return}, ${return_description}
/*/User Function GetOptX3(cCampo)
	Local aOpcoes := {}
	Local cTitulo := ""
	Local cOpcoes := &(ReadVar())
	Local lResult := .F.
	
	cRetGrupo := ""
	
	SX3->(DbSetOrder(2))//X3_CAMPO
	
	if(SX3->(DbSeek(cCampo)))
		aOpcoes := RetSx3Box(X3CBox(),,, 1)
		cTitulo := AllTrim(SX3->X3_TITULO)
		lResult := f_Opcoes( @cRetGrupo    ,;    //Variavel de Retorno
				              cTitulo     ,;    //Titulo da Coluna com as opcoes
				              @aOpcoes    ,;    //Opcoes de Escolha (Array de Opcoes)
				              @cOpcoes    ,;    //String de Opcoes para Retorno
				              NIL         ,;    //Nao Utilizado
				              NIL         ,;    //Nao Utilizado
				              .F.         ,;    //Se a Selecao sera de apenas 1 Elemento por vez
				              SX3->X3_TAMANHO,;    //Tamanho da Chave
				              Len(aOpcoes),;    //No maximo de elementos na variavel de retorno
				              .T.         ,;    //Inclui Botoes para Selecao de Multiplos Itens
				              .T.         ,;    //Se as opcoes serao montadas a partir de ComboBox de Campo ( X3_CBOX )
				              SX3->X3_CAMPO,;    //Qual o Campo para a Montagem do aOpcoes
				             .F.         ,;    //Nao Permite a Ordenacao
				             .F.         ,;    //Nao Permite a Pesquisa    
				             .F.         ,;    //Forca o Retorno Como Array
				             /*cF3*/          ;    //Consulta F3    
				            )
	endIf
	cRetGrupo := Left(cRetGrupo,Len(aOpcoes))
	
	SetMemVar(ReadVar(),cRetGrupo)	
Return lResult

/*/{Protheus.doc} Logger
	Funcao para padronizar mensagens de Log
@author PHILIPE.POMPEU
@since 29/08/2017
@param aLog, array, (Descrição do parâmetro)
@param cAssunto, character, (Descrição do parâmetro)
@param cTo, character, (Descrição do parâmetro)
@param bMail, booleano, (Descrição do parâmetro)
@param bShow, booleano, (Descrição do parâmetro)
@return ${return}, ${return_description}
/*/User Function Logger(aLog, cAssunto, cTo, bMail, bShow)
	Local cMensagem := ""
	Local nI := 0
	Local cCC := GETMV("MV_WFADMIN")
	Local aTemp := {}
	Default bMail := {|x|"<p>"+ x +"</p>"}
	Default bShow := {|x|{x}}
	
	cMensagem := "<p>O log abaixo é referente a <b>"+ cAssunto +"</b>, verificar.</p>"
	for nI:= 1 to Len(aLog)
		cMensagem += Eval(bMail,aLog[nI])
	next nI	
	
	if(IsBlind())
		U_SNDEMAIL(cAssunto,cMensagem, cTo, cCC,{},'')
	else
		If(MsgYesNo("Deseja enviar log por e-mail?"))
			U_SNDEMAIL(cAssunto,cMensagem, cTo, cCC,{},'')
		endIf
		
		for nI:= 1 to Len(aLog)
			aAdd(aTemp,Eval(bShow,aLog[nI]))
		next nI
		TmsMsgErr(aTemp)
		aSize(aTemp,0)
	endIf
	aSize(aLog,0)
Return nil

/*/{Protheus.doc} ExecFunc
	Avalia se deve usar loader para executar uma rotina
@author PHILIPE.POMPEU
@since 07/08/2017
@param bBlock, booleano, (Descrição do parâmetro)
@param cTexto, character, (Descrição do parâmetro)
@return ${return}, ${return_description}
/*/User Function ExecFunc(bBlock,cTexto)
	Local xResult := Nil
	Default cTexto := "Processando tarefa..."
	/*Realiza importacao*/
	if(IsBlind())
		xResult := Eval(bBlock)		
	else
		FWMsgRun( , bBlock, cTexto, "Processando..." ) 			
	endIf
Return xResult

/*/{Protheus.doc} ToExcel
	Rotina que exporta para o Excel uma consulta qualquer
@author PHILIPE.POMPEU
@since 07/08/2017
@param cTitulo, caractere, Título da Planilha
@param cQuery, caractere, Consulta SQL
@return ${return}, ${return_description}
/*/User Function ToExcel(cTitulo,cQuery)
	Local cMyAlias	:= GetNextAlias()
	Local aResult := {}
	Default cTitulo := "Exportação p/ Excel"
	
	aResult := ExcelQuery(cMyAlias, cQuery)	
	
	DlgToExcel({{"GETDADOS",cTitulo,aResult[1],aResult[2]}})
	
	aSize(aResult[2],0)
	aSize(aResult,0)
Return nil

Static Function ExcelQuery(cUmAlias,cQuery)
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
				cPict := SX3->X3_PICTURE
			else
				cPict := RetSx3Box(X3CBox(),,, 1)	
			endIf
			aAdd(aCab,{aStruct[nI,1],"C", SX3->X3_TAMANHO,SX3->X3_DECIMAL,SX3->X3_TITULO,cPict})
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