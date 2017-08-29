#INCLUDE 'PROTHEUS.CH'

User Function SNDEMAIL(clSubj,clBody,clTo,clCC,alAnexo, clNomDest)
	Local clServer 	:= AllTrim(GetNewPar("MV_RELSERV",""))
	Local clConta	:= AllTrim(GetNewPar("MV_RELACNT",""))
	Local clUser	:= AllTrim(GetNewPar("MV_RELACNT",""))
	Local clPass	:= AllTrim(GetNewPar("MV_RELPSW",""))
	Local llAut		:= GetNewPar("MV_RELAUTH",.F.)
	Local clFrom	:= AllTrim(GetNewPar("MV_RELFROM",""))
	                     
	ConOut("Enviando e-mail para " + clTo)
	               
	clSubj	:= StrTran(clSubj,"ã","a")
	clSubj	:= StrTran(clSubj,"á","a")
	clSubj	:= StrTran(clSubj,"ç","c")
	clSubj	:= StrTran(clSubj,"é","e")
	                         
	clSubj	:= StrTran(clSubj,"Ã","A")
	clSubj	:= StrTran(clSubj,"Á","A")
	clSubj	:= StrTran(clSubj,"Ç","C")
	clSubj	:= StrTran(clSubj,"É","E")
	
	
	llOk		:= MailSmtpOn( clServer, clConta, clPass )
	       
	clBody 		:= U_FBDYMAIL(clNomDest, clBody)
	
	If llOk
	
		If llAut
			llOk := MailAuth(clUser,clPass)
			
			If !llOk                                
				clUser := Left(clUser,At("@",clUser)-1)
				llOk := MailAuth(clUser,clPass)		
			EndIf                                   
			
		EndIf
		
		If llOk
			llOk :=	MailSend( clFrom, { clTo }, {    }, { clCC }, clSubj, clBody, alAnexo,.F. )
		Else     
			
			ConOut("Erro na autenticação de e-mail. : " + MAILGETERR())
		EndIf
	Else
		ConOut("Erro na conexão com o servidor de e-mail." + MAILGETERR())
	EndIf
	
	MailSmtpOff()

Return(.T.)

User Function FBDYMAIL(clNomDest, clBody, cDepart)
	Local cEnv := ''
	Default clNomDest := "" 
	Default cDepart := "Departamento de Compras"
	
	Do Case
		Case (AllTrim(cModulo) == 'COM')
			cDepart := "Departamento de Compras"
		Case (AllTrim(cModulo) == 'GCT')
			cDepart := "Gestão de Contratos"
		Case (AllTrim(cModulo) == 'FIN')
			cDepart := "Departamento Financeiro"
		Case (AllTrim(cModulo) == 'CTB')
			cDepart := "Departamento Contábil"
		OtherWise
			cDepart := "Tecnologia da Informação"		
	EndCase	
	
	clMsg := '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">'
	clMsg += '<HTML xmlns="http://www.w3.org/1999/xhtml">'
	clMsg += '<BODY style="background-color: #fff; color:#60707d;font-family:Open Sans;">'
	clMsg += '<div>'
	clMsg += '<img src="http://www.rvimola.com.br/wp-content/uploads/2016/02/rv_imola.png" title="RV ÍMOLA" alt="RV ÍMOLA" style="display: block;margin: 0 auto;">'
	clMsg += '<h3 style="color:#fff;background-color:#002063; padding:10px;text-transform: uppercase">Workflow '+ SM0->(M0_NOME) +'</h3>'
	clMsg += '</div>'	
	clMsg += '<p>Caro ' + clNomDest + ',</p>'
	clMsg += "<div>"
	clMsg += clBody
	clMsg += "</div>"
	clMsg += '<p>Atenciosamente,</p><p><b style="color:#158fd1">'+ cDepart + ' <a href="http://www.rvimola.com.br/">' + SM0->(M0_NOME)+ '</a></b></p>'
	
	cEnv := Upper(AllTrim(GetEnvServer()))
	
	if(cEnv != "PROD" .And. cEnv != "COMPILADOR" .And. cEnv != "WF")
		clMsg += '<h5 style="color:#fff;background-color:#FFA4AD; padding:10px;text-transform: uppercase">MENSAGEM DO AMBIENTE DE TESTES, DESCONSIDERAR.</h5>'	
	endIf
	
	clMsg += '<p style="text-align: center;">*** NÃO RESPONDER A ESSE E-MAIL, POIS TRATA-SE DE UMA MENSAGEM AUTOMÁTICA ***</p>'
	clMsg += '<hr/>'
	clMsg += '<p style="text-align: center;font-size:12px;">® RV ÍMOLA TODOS OS DIREITOS RESERVADOS</p>'
	clMsg += '</BODY></HTML>'
Return(clMsg)