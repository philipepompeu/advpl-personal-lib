#INCLUDE 'PROTHEUS.CH'

User Function SNDEMAIL(clSubj,clBody,clTo,clCC,alAnexo, clNomDest)
	Local clServer 	:= AllTrim(GetNewPar("MV_RELSERV",""))
	Local clConta	:= AllTrim(GetNewPar("MV_RELACNT",""))
	Local clUser	:= AllTrim(GetNewPar("MV_RELACNT",""))
	Local clPass	:= AllTrim(GetNewPar("MV_RELPSW",""))
	Local llAut		:= GetNewPar("MV_RELAUTH",.F.)
	Local clFrom	:= AllTrim(GetNewPar("MV_RELFROM",""))
	                     
	ConOut("Enviando e-mail para " + clTo)
	               
	clSubj	:= StrTran(clSubj,"�","a")
	clSubj	:= StrTran(clSubj,"�","a")
	clSubj	:= StrTran(clSubj,"�","c")
	clSubj	:= StrTran(clSubj,"�","e")
	                         
	clSubj	:= StrTran(clSubj,"�","A")
	clSubj	:= StrTran(clSubj,"�","A")
	clSubj	:= StrTran(clSubj,"�","C")
	clSubj	:= StrTran(clSubj,"�","E")
	
	
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
			
			ConOut("Erro na autentica��o de e-mail. : " + MAILGETERR())
		EndIf
	Else
		ConOut("Erro na conex�o com o servidor de e-mail." + MAILGETERR())
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
			cDepart := "Gest�o de Contratos"
		Case (AllTrim(cModulo) == 'FIN')
			cDepart := "Departamento Financeiro"
		Case (AllTrim(cModulo) == 'CTB')
			cDepart := "Departamento Cont�bil"
		OtherWise
			cDepart := "Tecnologia da Informa��o"		
	EndCase	
	
	clMsg := '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">'
	clMsg += '<HTML xmlns="http://www.w3.org/1999/xhtml">'
	clMsg += '<BODY style="background-color: #fff; color:#60707d;font-family:Open Sans;">'
	clMsg += '<div>'
	clMsg += '<img src="http://www.rvimola.com.br/wp-content/uploads/2016/02/rv_imola.png" title="RV �MOLA" alt="RV �MOLA" style="display: block;margin: 0 auto;">'
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
	
	clMsg += '<p style="text-align: center;">*** N�O RESPONDER A ESSE E-MAIL, POIS TRATA-SE DE UMA MENSAGEM AUTOM�TICA ***</p>'
	clMsg += '<hr/>'
	clMsg += '<p style="text-align: center;font-size:12px;">� RV �MOLA TODOS OS DIREITOS RESERVADOS</p>'
	clMsg += '</BODY></HTML>'
Return(clMsg)