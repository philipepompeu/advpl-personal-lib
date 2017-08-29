#INCLUDE 'TOTVS.CH'

CLASS TExcelRpt
	Data cTitulo as String
	Data cPergunte as String
	METHOD New(cPerg,cTitle) CONSTRUCTOR
	METHOD FixSX1()
	METHOD GetQuery()
	METHOD Execute()

ENDCLASS

//-----------------------------------------------------------------
METHOD New(cPerg,cTitle) CLASS TExcelRpt
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
	Local cQuery := ""
	Local cTitle := ""
	::FixSX1()
	if(Pergunte(::cPergunte))
		MakeSqlExpr(::cPergunte)
		cQuery := ::GetQuery()
		cTitle := ::cTitulo
		FWMsgRun( , {|x|U_ToExcel(cTitle,cQuery)}, "Gerando relatório", "Processando..." )		
	endIf
	RestArea(aArea)
Return