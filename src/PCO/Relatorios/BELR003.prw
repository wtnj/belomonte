#INCLUDE "Protheus.ch"
#INCLUDE "TopConn.ch"

#DEFINE c_ent CHR(13)+CHR(10)

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³BELR003   ºAutor  ³Marcelo Amaral      º Data ³  06/02/15   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Programa para gerar planilha de movimentações conforme     º±±
±±º          ³ cubo gerencial selecionado                                 º±± 
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP                                                         º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

User Function BELR003()  

Local _aArea := GetArea() 
Local cPerg := Padr("BELR003",Len(SX1->X1_GRUPO))      	             
Private cDesc1 := ""
Private cDesc2 := ""

PutSx1(cPerg,"01","Cubo Gerencial ?"	,"","","mv_ch1","C",TAMSX3("AL1_CONFIG")[1]	,0,0,"G","","AL1"	,"","","MV_PAR01","","","","","","","","","","","","","","","","","","","","","","","")
PutSx1(cPerg,"02","Data de ?"			,"","","mv_ch2","D",TAMSX3("AKD_DATA")[1]  	,0,0,"G","",""		,"","","MV_PAR02","","","","","","","","","","","","","","","","","","","","","","","")
PutSx1(cPerg,"03","Data Ate ?"			,"","","mv_ch3","D",TAMSX3("AKD_DATA")[1]  	,0,0,"G","",""		,"","","MV_PAR03","","","","","","","","","","","","","","","","","","","","","","","")
PutSx1(cPerg,"04","Config Cubo 1 ?"		,"","","mv_ch4","C",TAMSX3("AL3_CODIGO")[1]	,0,0,"G","","AL3"	,"","","MV_PAR04","","","","","","","","","","","","","","","","","","","","","","","")
PutSx1(cPerg,"05","Config Cubo 2 ?"		,"","","mv_ch5","C",TAMSX3("AL3_CODIGO")[1]	,0,0,"G","","AL3"	,"","","MV_PAR05","","","","","","","","","","","","","","","","","","","","","","","")
PutSx1(cPerg,"06","C.Custos de ?"		,"","","mv_ch6","C",TAMSX3("CTT_CUSTO")[1]	,0,0,"G","","CTT"	,"","","MV_PAR06","","","","","","","","","","","","","","","","","","","","","","","")
PutSx1(cPerg,"07","C.Custos Ate ?"		,"","","mv_ch7","C",TAMSX3("CTT_CUSTO")[1]	,0,0,"G","","CTT"	,"","","MV_PAR07","","","","","","","","","","","","","","","","","","","","","","","")

If !Pergunte(cPerg,.T.)
	RestArea(_aArea)
	Return
endif

DbSelectArea("AKW")
AKW->(DbSetOrder(1))
If !AKW->(DbSeek(xFilial("AKW")+mv_par01))
	MsgStop("Entidades do Cubo Não Cadastradas!")
	RestArea(_aArea)
	Return
Endif

dbSelectArea("AL3")
AL3->(dbSetOrder(3))
If !AL3->(DbSeek(xFilial("AL3")+mv_par01+mv_par04))
	MsgStop("Configuração do Cubo 1 Não Cadastrada!")
	RestArea(_aArea)
	Return
endif

dbSelectArea("AL4")
AL4->(dbSetOrder(3))
If !AL4->(DbSeek(xFilial("AL4")+mv_par01+mv_par04))
	MsgStop("Configuração do Cubo 1 Não Cadastrada!")
	RestArea(_aArea)
	Return
endif

cFiltro1 := Alltrim(AL4->AL4_FILTER)

cFilQry1 := Alltrim(StrTran(StrTran(StrTran(StrTran(StrTran(AL4->AL4_FILTER,"==","="),'"',"'"),".OR.","OR"),".AND.","AND"),chr(13)+chr(10),""))

cDesc1 := Upper(AL3->AL3_DESCRI)

dbSelectArea("AL3")
AL3->(dbSetOrder(3))
If !AL3->(DbSeek(xFilial("AL3")+mv_par01+mv_par05))
	MsgStop("Configuração do Cubo 2 Não Cadastrada!")
	RestArea(_aArea)
	Return
endif

dbSelectArea("AL4")
AL4->(dbSetOrder(3))
If !AL4->(DbSeek(xFilial("AL4")+mv_par01+mv_par05))
	MsgStop("Configuração do Cubo 2 Não Cadastrada!")
	RestArea(_aArea)
	Return
endif

cFiltro2 := Alltrim(AL4->AL4_FILTER)

cFilQry2 := Alltrim(StrTran(StrTran(StrTran(StrTran(StrTran(AL4->AL4_FILTER,"==","="),'"',"'"),".OR.","OR"),".AND.","AND"),chr(13)+chr(10),""))

cDesc2 := Upper(AL3->AL3_DESCRI)

Processa({|| BELR003A()},"Aguarde...")
	
RestArea(_aArea)
                                               
Return

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³BELR003A   ºAutor  ³Marcelo Amaral     º Data ³  06/02/15   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Função para gerar a planilha de excel                      º±±
±±º          ³                                                            º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP                                                         º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function BELR003A()

Local _cFile := "BELR003.xml"                      
Local _cPath := GetTempPath() //Pega a Pasta de Arquivo temporário                 
Local cQuery := ""
Private cAlias := GetNextAlias()

Private _nHandle
Private aBreaks := {}
                      
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³Monta o Nome do Arquivo³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
_cArquivo := Alltrim(_cPath)+iif(Right(Alltrim(_cPath),1)=="\","","\")+Alltrim(_cFile)
If File( _cArquivo )
	FErase( _cArquivo )
End If

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³Cria o Arquivo Fisico³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
_nHandle := MsFCreate( _cArquivo )     

//nQtdEntid := CtbQtdEntd()
nCont := 0

//Monta vetor de configurações do cubo selecionado - com chaves, relacionamentos e quebras

cChvRel := ""
While !AKW->(Eof()) .And. AKW->AKW_FILIAL == xFilial("AKW") .and. AKW->AKW_COD == mv_par01
	cChvRel += iif(cChvRel <> "","+","") + Alltrim(StrTran(AKW->AKW_RELAC,Alltrim(AKW->AKW_ALIAS)+"->",""))
	nCont += 1
	aadd(aBreaks,)
	aBreaks[Len(aBreaks)] := {Alltrim(AKW->AKW_DESCRI),Alltrim(StrTran(AKW->AKW_CHAVER,"AKD->","")),Alltrim(StrTran(AKW->AKW_CONCCH,"AKD->","")),Alltrim(AKW->AKW_ALIAS),Alltrim(StrTran(AKW->AKW_RELAC,Alltrim(AKW->AKW_ALIAS)+"->","")),Alltrim(StrTran(AKW->AKW_DESCRE,Alltrim(AKW->AKW_ALIAS)+"->","")),cChvRel,AKW->AKW_NIVEL}
	cOrder := AllTrim(AKW->AKW_CONCCH)
	AKW->(DbSkip())
Enddo  

nQtdEntid := Len(aBreaks)

cTab := aBreaks[nQtdEntid,4]
cCpo := aBreaks[nQtdEntid,2]

cOrder  := StrTran(cOrder,"AKD->","")  //primeiro tiro o ALIAS
cOrder  := StrTran(cOrder,"+",",")     //depois mudo o operador soma (concatenar) para virgula

//Monta query dinamica de acordo com o cubo selecionado

cQuery := ""
cQuery += "SELECT AKD_CC AS CENCUS,AK5_COSUP,"+cTab+".R_E_C_N_O_ AS REC"+cTab+",AKD_DATA,(CASE WHEN AKD_TIPO = 2 THEN (AKD_VALOR1*(-1)) ELSE AKD_VALOR1 END) AS AKD_VALOR1,"+cOrder+","
For nCont := 1 to Len(aBreaks)
	cQuery += aBreaks[nCont,5]+","+aBreaks[nCont,6] + iif(nCont < Len(aBreaks),",","")
Next
cQuery += " " + c_ent

cQuery += "FROM "+RetSqlName("AK5")+" AK5 " + c_ent
cQuery += "INNER JOIN "+RetSqlName("CTD")+" CTD " + c_ent
cQuery += "ON AK5_FILIAL = '"+xFilial("AK5")+"' " + c_ent
cQuery += "AND CTD_FILIAL = '"+xFilial("CTD")+"' " + c_ent
cQuery += "AND CTD_XCO = AK5_CODIGO " + c_ent
cQuery += "AND AK5.D_E_L_E_T_ = ' ' " + c_ent
cQuery += "AND CTD.D_E_L_E_T_ = ' ' " + c_ent
cQuery += "LEFT OUTER JOIN AKD010 AKD " + c_ent
cQuery += "ON AKD_CO = AK5_CODIGO " + c_ent
cQuery += "AND AKD_ITCTB = CTD_ITEM " + c_ent
cQuery += "AND AKD.D_E_L_E_T_ = ' ' " + c_ent

For nCont := 3 to Len(aBreaks)
	cQuery += "LEFT OUTER JOIN "+RetSqlName(aBreaks[nCont,4])+" "+aBreaks[nCont,4]+" " + c_ent
	cQuery += "ON "+iif(Substr(aBreaks[nCont,4],1,1)=="S",Substr(aBreaks[nCont,4],2,2),aBreaks[nCont,4])+"_FILIAL = '"+xfilial(aBreaks[nCont,4])+"' " + c_ent
	cQuery += "AND "+aBreaks[nCont,5]+" = "+aBreaks[nCont,2]+" " + c_ent
	cQuery += "AND "+aBreaks[nCont,4]+".D_E_L_E_T_ = ' ' " + c_ent
Next

cQuery += "WHERE (AKD_FILIAL = '"+xfilial("AKD")+"' OR AKD_FILIAL IS NULL) " + c_ent
cQuery += "AND (AKD_STATUS = '1' OR AKD_STATUS IS NULL) " + c_ent
cQuery += "AND ((AKD_DATA BETWEEN '"+DTOS(MV_PAR02)+"' AND '"+DTOS(MV_PAR03)+"') OR AKD_DATA IS NULL) "+ c_ent      
cQuery += "AND (AKD.D_E_L_E_T_ = ' ' OR AKD.D_E_L_E_T_ IS NULL) " + c_ent
//cQuery += "AND ((AKD_CC >= '"+mv_par06+"' AND AKD_CC <= '"+mv_par07+"') OR AKD_CC IS NULL) " + c_ent

cQuery += "AND ("+cCpo+" IS NULL OR (" + cFilQry1 + " " + c_ent
cQuery += "OR " + cFilQry2 + ")) " + c_ent

//inicio

cQuery += "UNION ALL " + c_ent

cQuery += "SELECT AKD_CC AS CENCUS,AK5_COSUP,"+cTab+".R_E_C_N_O_ AS REC"+cTab+",AKD_DATA,(CASE WHEN AKD_TIPO = 2 THEN (AKD_VALOR1*(-1)) ELSE AKD_VALOR1 END) AS AKD_VALOR1,"+cOrder+","
For nCont := 1 to Len(aBreaks)
	cQuery += aBreaks[nCont,5]+","+aBreaks[nCont,6] + iif(nCont < Len(aBreaks),",","")
Next
cQuery += " " + c_ent

cQuery += "FROM "+RetSqlName("AK5")+" AK5 " + c_ent

cQuery += "LEFT OUTER JOIN AKD010 AKD " + c_ent
cQuery += "ON AKD_CO = AK5_CODIGO " + c_ent
cQuery += "AND AKD.D_E_L_E_T_ = ' ' " + c_ent

For nCont := 2 to Len(aBreaks)
	cQuery += "LEFT OUTER JOIN "+RetSqlName(aBreaks[nCont,4])+" "+aBreaks[nCont,4]+" " + c_ent
	cQuery += "ON "+iif(Substr(aBreaks[nCont,4],1,1)=="S",Substr(aBreaks[nCont,4],2,2),aBreaks[nCont,4])+"_FILIAL = '"+xfilial(aBreaks[nCont,4])+"' " + c_ent
	cQuery += "AND "+aBreaks[nCont,5]+" = "+aBreaks[nCont,2]+" " + c_ent
	cQuery += "AND "+aBreaks[nCont,4]+".D_E_L_E_T_ = ' ' " + c_ent
Next

cQuery += "WHERE AK5_FILIAL = '"+xFilial("AK5")+"' " + c_ent
cQuery += "AND AK5_TIPO = '2' " + c_ent
cQuery += "AND AK5.D_E_L_E_T_ = ' ' " + c_ent
cQuery += "AND AK5_CODIGO NOT IN ( " + c_ent

cQuery += "SELECT DISTINCT AK5_CODIGO " + c_ent
cQuery += "FROM "+RetSqlName("AK5")+" AK5_2 " + c_ent
cQuery += "INNER JOIN "+RetSqlName("CTD")+" CTD_2 " + c_ent
cQuery += "ON AK5_2.AK5_FILIAL = '"+xFilial("AK5")+"' " + c_ent
cQuery += "AND CTD_2.CTD_FILIAL = '"+xFilial("CTD")+"' " + c_ent
cQuery += "AND CTD_2.CTD_XCO = AK5_2.AK5_CODIGO " + c_ent
cQuery += "AND AK5_2.D_E_L_E_T_ = ' ' " + c_ent
cQuery += "AND CTD_2.D_E_L_E_T_ = ' ') " + c_ent

cQuery += "AND (AKD_FILIAL = '"+xfilial("AKD")+"' OR AKD_FILIAL IS NULL) " + c_ent
cQuery += "AND (AKD_STATUS = '1' OR AKD_STATUS IS NULL) " + c_ent
cQuery += "AND ((AKD_DATA BETWEEN '"+DTOS(MV_PAR02)+"' AND '"+DTOS(MV_PAR03)+"') OR AKD_DATA IS NULL) "+ c_ent      
cQuery += "AND (AKD.D_E_L_E_T_ = ' ' OR AKD.D_E_L_E_T_ IS NULL) " + c_ent
//cQuery += "AND ((AKD_CC >= '"+mv_par06+"' AND AKD_CC <= '"+mv_par07+"') OR AKD_CC IS NULL) " + c_ent

cQuery += "AND ("+cCpo+" IS NULL OR (" + cFilQry1 + " " + c_ent
cQuery += "OR " + cFilQry2 + ")) " + c_ent

//fim

cOrder  := aBreaks[1,5]+","+aBreaks[2,5]+","+cOrder

cQuery += "ORDER BY "+cOrder + c_ent

MemoWrite("c:\temp\",cQuery)

If Select(cAlias) > 0
	dbSelectArea(cAlias)
	(cAlias)->(dbCloseArea())
EndIf

dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAlias,.T.,.T.)       

dbSelectArea(cAlias)

nTotal01 := 0
nTotal02 := 0
nTotal03 := 0
nTotal04 := 0
nTotal05 := 0
nTotal06 := 0
nTotal07 := 0
nTotal08 := 0
nTotal09 := 0
nTotal10 := 0
nTotal11 := 0
nTotal12 := 0
nTotal13 := 0
nTotal14 := 0
nTotal15 := 0
nTotal16 := 0
nTotal17 := 0
nTotal18 := 0
nTotal19 := 0
nTotal20 := 0
nTotal21 := 0
nTotal22 := 0
nTotal23 := 0
nTotal24 := 0
nTotal25 := 0
nTotal26 := 0

ProcRegua((cAlias)->(RecCount()))

//Monta vetor que vai alimentar a planilha excel
aLinha := {}
While !(cAlias)->(EOF())
	IncProc("Buscando Informações...")
	For nCont := 1 to Len(aBreaks) - 1
		if aBreaks[nCont,8] == "01"
			aCoSup := {}
			cCoSup := Alltrim((cAlias)->AK5_COSUP)
			While !(cCoSup == "")
				aadd(aCoSup,cCoSup)
				dbSelectArea("AK5")
				AK5->(dbSetOrder(1))
				AK5->(dbSeek(xFilial("AK5")+cCoSup))
				cCoSup := Alltrim(AK5->AK5_COSUP)
			End
			aSort(aCoSup)
			For nCont3 := 1 to Len(aCoSup)
				nLinha := Ascan(aLinha,{|x| Alltrim(x[1]) == Alltrim(aBreaks[nCont,8]) .and. ;
					Alltrim(x[2]) == Alltrim(aCoSup[nCont3])+;
					Alltrim(aCoSup[nCont3])})
				if nLinha = 0
					aadd(aLinha,{Alltrim(aBreaks[nCont,8]),;
						Alltrim(aCoSup[nCont3])+;
						Alltrim(aCoSup[nCont3]),;
						Alltrim(aCoSup[nCont3]),;
						Alltrim(&("(cAlias)->"+aBreaks[nCont,6])),0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0})
					nLinha := Len(aLinha)
				endif
				if !(Alltrim((cAlias)->AKD_DATA) == "")
					dbSelectArea(cTab)
					&(cTab+"->(dbGoTo((cAlias)->REC"+cTab+"))")
					if &(cFiltro1)
						if (cAlias)->CENCUS >= mv_par06 .and. (cAlias)->CENCUS <= mv_par07
							aLinha[nLinha,((Val(Substr((cAlias)->AKD_DATA,5,2))*2)-1)+4] += (cAlias)->AKD_VALOR1
							aLinha[nLinha,29] += (cAlias)->AKD_VALOR1
							if nCont3 = 1
								&("nTotal"+StrZero(((Val(Substr((cAlias)->AKD_DATA,5,2))*2)-1),2)) += (cAlias)->AKD_VALOR1
								nTotal25 += (cAlias)->AKD_VALOR1
							endif
						endif
					elseif &(cFiltro2)
						if (cAlias)->CENCUS >= mv_par06 .and. (cAlias)->CENCUS <= mv_par07
							aLinha[nLinha,(Val(Substr((cAlias)->AKD_DATA,5,2))*2)+4] += (cAlias)->AKD_VALOR1
							aLinha[nLinha,30] += (cAlias)->AKD_VALOR1
							if nCont3 = 1
								&("nTotal"+StrZero((Val(Substr((cAlias)->AKD_DATA,5,2))*2),2)) += (cAlias)->AKD_VALOR1
								nTotal26 += (cAlias)->AKD_VALOR1
							endif
						endif
					endif
				endif
			Next	
		endif
		lOutItm := .F.
		if aBreaks[nCont,4] == "CTD"
			cOutItm := GetMV("MV_XOUTITM")
			if !(cOutItm == "")
				if Alltrim(&("(cAlias)->"+aBreaks[nCont,5])) $ cOutItm
					lOutItm := .T.
				endif
			endif
		endif
		if !lOutItm
			nLinha := Ascan(aLinha,{|x| Alltrim(x[1]) == Alltrim(aBreaks[nCont,8]) .and. ;
				Alltrim(x[2]) == iif(Alltrim(&("(cAlias)->("+aBreaks[nCont,3]+")"))<>"",;
				Alltrim(&("(cAlias)->("+aBreaks[nCont,3]+")")),;
				Alltrim(&("(cAlias)->("+aBreaks[nCont,7]+")")))+;
				iif(Alltrim(&("(cAlias)->"+aBreaks[nCont,2]))<>"",;
				Alltrim(&("(cAlias)->"+aBreaks[nCont,2])),;
				Alltrim(&("(cAlias)->"+aBreaks[nCont,5])))})
			if nLinha = 0
				aadd(aLinha,{Alltrim(aBreaks[nCont,8]),;
					iif(Alltrim(&("(cAlias)->("+aBreaks[nCont,3]+")"))<>"",;
					Alltrim(&("(cAlias)->("+aBreaks[nCont,3]+")")),;
					Alltrim(&("(cAlias)->("+aBreaks[nCont,7]+")")))+;
					iif(Alltrim(&("(cAlias)->"+aBreaks[nCont,2]))<>"",;
					Alltrim(&("(cAlias)->"+aBreaks[nCont,2])),;
					Alltrim(&("(cAlias)->"+aBreaks[nCont,5]))),;
					iif(Alltrim(&("(cAlias)->"+aBreaks[nCont,2]))<>"",;
					Alltrim(&("(cAlias)->"+aBreaks[nCont,2])),;
					Alltrim(&("(cAlias)->"+aBreaks[nCont,5]))),;
					Alltrim(&("(cAlias)->"+aBreaks[nCont,6])),0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0})
				nLinha := Len(aLinha)
			endif
			if !(Alltrim((cAlias)->AKD_DATA) == "")
				dbSelectArea(cTab)
				&(cTab+"->(dbGoTo((cAlias)->REC"+cTab+"))")
				if &(cFiltro1)
					if (cAlias)->CENCUS >= mv_par06 .and. (cAlias)->CENCUS <= mv_par07
						aLinha[nLinha,((Val(Substr((cAlias)->AKD_DATA,5,2))*2)-1)+4] += (cAlias)->AKD_VALOR1
						aLinha[nLinha,29] += (cAlias)->AKD_VALOR1
					endif
				elseif &(cFiltro2)
					if (cAlias)->CENCUS >= mv_par06 .and. (cAlias)->CENCUS <= mv_par07
						aLinha[nLinha,(Val(Substr((cAlias)->AKD_DATA,5,2))*2)+4] += (cAlias)->AKD_VALOR1
						aLinha[nLinha,30] += (cAlias)->AKD_VALOR1
					endif
				endif
			endif
		endif	
	Next
	(cAlias)->(dbSkip())
End

fWrite(_nHandle,'<?xml version="1.0"?>' + c_ent)
fWrite(_nHandle,'<?mso-application progid="Excel.Sheet"?>' + c_ent)
fWrite(_nHandle,'<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"' + c_ent)
fWrite(_nHandle,' xmlns:o="urn:schemas-microsoft-com:office:office"' + c_ent)
fWrite(_nHandle,' xmlns:x="urn:schemas-microsoft-com:office:excel"' + c_ent)
fWrite(_nHandle,' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"' + c_ent)
fWrite(_nHandle,' xmlns:html="http://www.w3.org/TR/REC-html40">' + c_ent)
fWrite(_nHandle,' <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">' + c_ent)
fWrite(_nHandle,'  <Author>PROTHEUS 11</Author>' + c_ent)
fWrite(_nHandle,'  <LastAuthor>PROTHEUS 11</LastAuthor>' + c_ent)
fWrite(_nHandle,'  <Created>2015-02-04T14:04:48Z</Created>' + c_ent)
fWrite(_nHandle,'  <LastSaved>2015-02-15T15:04:18Z</LastSaved>' + c_ent)
fWrite(_nHandle,'  <Version>15.00</Version>' + c_ent)
fWrite(_nHandle,' </DocumentProperties>' + c_ent)
fWrite(_nHandle,' <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">' + c_ent)
fWrite(_nHandle,'  <AllowPNG/>' + c_ent)
fWrite(_nHandle,' </OfficeDocumentSettings>' + c_ent)
fWrite(_nHandle,' <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">' + c_ent)
fWrite(_nHandle,'  <WindowHeight>11535</WindowHeight>' + c_ent)
fWrite(_nHandle,'  <WindowWidth>25200</WindowWidth>' + c_ent)
fWrite(_nHandle,'  <WindowTopX>0</WindowTopX>' + c_ent)
fWrite(_nHandle,'  <WindowTopY>0</WindowTopY>' + c_ent)
fWrite(_nHandle,'  <ProtectStructure>False</ProtectStructure>' + c_ent)
fWrite(_nHandle,'  <ProtectWindows>False</ProtectWindows>' + c_ent)
fWrite(_nHandle,' </ExcelWorkbook>' + c_ent)

fWrite(_nHandle,' <Styles>' + c_ent)
fWrite(_nHandle,'  <Style ss:ID="Default" ss:Name="Normal">' + c_ent)
fWrite(_nHandle,'   <Alignment ss:Vertical="Bottom"/>' + c_ent)
fWrite(_nHandle,'   <Borders/>' + c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>' + c_ent)
fWrite(_nHandle,'   <Interior/>' + c_ent)
fWrite(_nHandle,'   <NumberFormat/>' + c_ent)
fWrite(_nHandle,'   <Protection/>' + c_ent)
fWrite(_nHandle,'  </Style>' + c_ent)
fWrite(_nHandle,'  <Style ss:ID="s16" ss:Name="Normal 9">' + c_ent)
fWrite(_nHandle,'   <Alignment ss:Vertical="Bottom"/>' + c_ent)
fWrite(_nHandle,'   <Borders/>' + c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>' + c_ent)
fWrite(_nHandle,'   <Interior/>' + c_ent)
fWrite(_nHandle,'   <NumberFormat ss:Format="[ENG][$-409]mmm\-yy;@"/>' + c_ent)
fWrite(_nHandle,'   <Protection/>' + c_ent)
fWrite(_nHandle,'  </Style>' + c_ent)
fWrite(_nHandle,'  <Style ss:ID="s17" ss:Name="Vírgula 4">' + c_ent)
fWrite(_nHandle,'   <NumberFormat ss:Format="_(* #,##0.00_);_(* \(#,##0.00\);_(* &quot;-&quot;??_);_(@_)"/>' + c_ent)
fWrite(_nHandle,'  </Style>' + c_ent)
fWrite(_nHandle,'  <Style ss:ID="m300752760" ss:Parent="s16">' + c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>' + c_ent)
fWrite(_nHandle,'   <Borders>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Bottom" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Left" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Right" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Top" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'   </Borders>' + c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="9" ss:Color="#FFFFFF"' + c_ent)
fWrite(_nHandle,'    ss:Bold="1"/>' + c_ent)
fWrite(_nHandle,'   <Interior ss:Color="#222B35" ss:Pattern="Solid"/>' + c_ent)
fWrite(_nHandle,'   <NumberFormat ss:Format="[ENG][$-409]d\-mmm;@"/>' + c_ent)
fWrite(_nHandle,'  </Style>' + c_ent)
fWrite(_nHandle,'  <Style ss:ID="m300752780" ss:Parent="s16">' + c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>' + c_ent)
fWrite(_nHandle,'   <Borders>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Bottom" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Left" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Right" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Top" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'   </Borders>' + c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="9" ss:Color="#FFFFFF"' + c_ent)
fWrite(_nHandle,'    ss:Bold="1"/>' + c_ent)
fWrite(_nHandle,'   <Interior ss:Color="#222B35" ss:Pattern="Solid"/>' + c_ent)
fWrite(_nHandle,'   <NumberFormat ss:Format="[ENG][$-409]d\-mmm;@"/>' + c_ent)
fWrite(_nHandle,'  </Style>' + c_ent)
fWrite(_nHandle,'  <Style ss:ID="m300752800" ss:Parent="s16">' + c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>' + c_ent)
fWrite(_nHandle,'   <Borders>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Bottom" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Left" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Right" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Top" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'   </Borders>' + c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="9" ss:Color="#FFFFFF"' + c_ent)
fWrite(_nHandle,'    ss:Bold="1"/>' + c_ent)
fWrite(_nHandle,'   <Interior ss:Color="#222B35" ss:Pattern="Solid"/>' + c_ent)
fWrite(_nHandle,'   <NumberFormat ss:Format="[ENG][$-409]d\-mmm;@"/>' + c_ent)
fWrite(_nHandle,'  </Style>' + c_ent)
fWrite(_nHandle,'  <Style ss:ID="m300752136" ss:Parent="s16">' + c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>' + c_ent)
fWrite(_nHandle,'   <Borders>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Bottom" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Left" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Right" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Top" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'   </Borders>' + c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="9" ss:Color="#FFFFFF"' + c_ent)
fWrite(_nHandle,'    ss:Bold="1"/>' + c_ent)
fWrite(_nHandle,'   <Interior ss:Color="#222B35" ss:Pattern="Solid"/>' + c_ent)
fWrite(_nHandle,'   <NumberFormat ss:Format="[ENG][$-409]d\-mmm;@"/>' + c_ent)
fWrite(_nHandle,'  </Style>' + c_ent)
fWrite(_nHandle,'  <Style ss:ID="m300752156" ss:Parent="s16">' + c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>' + c_ent)
fWrite(_nHandle,'   <Borders>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Bottom" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Left" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Right" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Top" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'   </Borders>' + c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="9" ss:Color="#FFFFFF"' + c_ent)
fWrite(_nHandle,'    ss:Bold="1"/>' + c_ent)
fWrite(_nHandle,'   <Interior ss:Color="#222B35" ss:Pattern="Solid"/>' + c_ent)
fWrite(_nHandle,'   <NumberFormat ss:Format="[ENG][$-409]d\-mmm;@"/>' + c_ent)
fWrite(_nHandle,'  </Style>' + c_ent)
fWrite(_nHandle,'  <Style ss:ID="m300752176" ss:Parent="s16">' + c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>' + c_ent)
fWrite(_nHandle,'   <Borders>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Bottom" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Left" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Right" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Top" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'   </Borders>' + c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="9" ss:Color="#FFFFFF"' + c_ent)
fWrite(_nHandle,'    ss:Bold="1"/>' + c_ent)
fWrite(_nHandle,'   <Interior ss:Color="#222B35" ss:Pattern="Solid"/>' + c_ent)
fWrite(_nHandle,'   <NumberFormat ss:Format="[ENG][$-409]d\-mmm;@"/>' + c_ent)
fWrite(_nHandle,'  </Style>' + c_ent)
fWrite(_nHandle,'  <Style ss:ID="m300752196" ss:Parent="s16">' + c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>' + c_ent)
fWrite(_nHandle,'   <Borders>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Bottom" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Left" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Right" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Top" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'   </Borders>' + c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="9" ss:Color="#FFFFFF"' + c_ent)
fWrite(_nHandle,'    ss:Bold="1"/>' + c_ent)
fWrite(_nHandle,'   <Interior ss:Color="#222B35" ss:Pattern="Solid"/>' + c_ent)
fWrite(_nHandle,'   <NumberFormat ss:Format="[ENG][$-409]d\-mmm;@"/>' + c_ent)
fWrite(_nHandle,'  </Style>' + c_ent)
fWrite(_nHandle,'  <Style ss:ID="m300752216" ss:Parent="s16">' + c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>' + c_ent)
fWrite(_nHandle,'   <Borders>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Bottom" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Left" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Right" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Top" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'   </Borders>' + c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="9" ss:Color="#FFFFFF"' + c_ent)
fWrite(_nHandle,'    ss:Bold="1"/>' + c_ent)
fWrite(_nHandle,'   <Interior ss:Color="#222B35" ss:Pattern="Solid"/>' + c_ent)
fWrite(_nHandle,'   <NumberFormat ss:Format="[ENG][$-409]d\-mmm;@"/>' + c_ent)
fWrite(_nHandle,'  </Style>' + c_ent)
fWrite(_nHandle,'  <Style ss:ID="m300752236" ss:Parent="s16">' + c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>' + c_ent)
fWrite(_nHandle,'   <Borders>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Bottom" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Left" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Right" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Top" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'   </Borders>' + c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="9" ss:Color="#FFFFFF"' + c_ent)
fWrite(_nHandle,'    ss:Bold="1"/>' + c_ent)
fWrite(_nHandle,'   <Interior ss:Color="#222B35" ss:Pattern="Solid"/>' + c_ent)
fWrite(_nHandle,'   <NumberFormat ss:Format="[ENG][$-409]d\-mmm;@"/>' + c_ent)
fWrite(_nHandle,'  </Style>' + c_ent)
fWrite(_nHandle,'  <Style ss:ID="m300752256" ss:Parent="s16">' + c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>' + c_ent)
fWrite(_nHandle,'   <Borders>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Bottom" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Left" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Right" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Top" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'   </Borders>' + c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="9" ss:Color="#FFFFFF"' + c_ent)
fWrite(_nHandle,'    ss:Bold="1"/>' + c_ent)
fWrite(_nHandle,'   <Interior ss:Color="#222B35" ss:Pattern="Solid"/>' + c_ent)
fWrite(_nHandle,'   <NumberFormat ss:Format="[ENG][$-409]d\-mmm;@"/>' + c_ent)
fWrite(_nHandle,'  </Style>' + c_ent)
fWrite(_nHandle,'  <Style ss:ID="m300752276" ss:Parent="s16">' + c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>' + c_ent)
fWrite(_nHandle,'   <Borders>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Bottom" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Left" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Right" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Top" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'   </Borders>' + c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="9" ss:Color="#FFFFFF"' + c_ent)
fWrite(_nHandle,'    ss:Bold="1"/>' + c_ent)
fWrite(_nHandle,'   <Interior ss:Color="#222B35" ss:Pattern="Solid"/>' + c_ent)
fWrite(_nHandle,'   <NumberFormat ss:Format="[ENG][$-409]d\-mmm;@"/>' + c_ent)
fWrite(_nHandle,'  </Style>' + c_ent)
fWrite(_nHandle,'  <Style ss:ID="m300752296" ss:Parent="s16">' + c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>' + c_ent)
fWrite(_nHandle,'   <Borders>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Bottom" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Left" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Right" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Top" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'   </Borders>' + c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="9" ss:Color="#FFFFFF"' + c_ent)
fWrite(_nHandle,'    ss:Bold="1"/>' + c_ent)
fWrite(_nHandle,'   <Interior ss:Color="#222B35" ss:Pattern="Solid"/>' + c_ent)
fWrite(_nHandle,'   <NumberFormat ss:Format="[ENG][$-409]d\-mmm;@"/>' + c_ent)
fWrite(_nHandle,'  </Style>' + c_ent)
fWrite(_nHandle,'  <Style ss:ID="m300752316" ss:Parent="s16">' + c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>' + c_ent)
fWrite(_nHandle,'   <Borders>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Bottom" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Left" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Right" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Top" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'   </Borders>' + c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="9" ss:Color="#FFFFFF"' + c_ent)
fWrite(_nHandle,'    ss:Bold="1"/>' + c_ent)
fWrite(_nHandle,'   <Interior ss:Color="#222B35" ss:Pattern="Solid"/>' + c_ent)
fWrite(_nHandle,'   <NumberFormat ss:Format="[ENG][$-409]d\-mmm;@"/>' + c_ent)
fWrite(_nHandle,'  </Style>' + c_ent)
fWrite(_nHandle,'  <Style ss:ID="s19" ss:Parent="s16">' + c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="9" ss:Color="#000000"/>' + c_ent)
fWrite(_nHandle,'  </Style>' + c_ent)
fWrite(_nHandle,'  <Style ss:ID="s20" ss:Parent="s16">' + c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="9" ss:Color="#000000"/>' + c_ent)
fWrite(_nHandle,'   <Interior/>' + c_ent)
fWrite(_nHandle,'  </Style>' + c_ent)
fWrite(_nHandle,'  <Style ss:ID="s21" ss:Parent="s16">' + c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>' + c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="9" ss:Bold="1"/>' + c_ent)
fWrite(_nHandle,'   <NumberFormat ss:Format="Standard"/>' + c_ent)
fWrite(_nHandle,'  </Style>' + c_ent)
fWrite(_nHandle,'  <Style ss:ID="s22" ss:Parent="s16">' + c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>' + c_ent)
fWrite(_nHandle,'   <Borders/>' + c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="9" ss:Color="#000000"' + c_ent)
fWrite(_nHandle,'    ss:Bold="1"/>' + c_ent)
fWrite(_nHandle,'   <Interior/>' + c_ent)
fWrite(_nHandle,'   <NumberFormat ss:Format="Standard"/>' + c_ent)
fWrite(_nHandle,'  </Style>' + c_ent)
fWrite(_nHandle,'  <Style ss:ID="s24" ss:Parent="s16">' + c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>' + c_ent)
fWrite(_nHandle,'   <Borders>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Bottom" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Left" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Right" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Top" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'   </Borders>' + c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="9" ss:Color="#000000"' + c_ent)
fWrite(_nHandle,'    ss:Bold="1"/>' + c_ent)
fWrite(_nHandle,'   <Interior/>' + c_ent)
fWrite(_nHandle,'   <NumberFormat ss:Format="[ENG][$-409]d\-mmm;@"/>' + c_ent)
fWrite(_nHandle,'  </Style>' + c_ent)
fWrite(_nHandle,'  <Style ss:ID="s26" ss:Parent="s16">' + c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>' + c_ent)
fWrite(_nHandle,'   <Borders>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Bottom" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Left" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Right" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Top" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'   </Borders>' + c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="9" ss:Color="#FFFFFF"' + c_ent)
fWrite(_nHandle,'    ss:Bold="1"/>' + c_ent)
fWrite(_nHandle,'   <Interior ss:Color="#222B35" ss:Pattern="Solid"/>' + c_ent)
fWrite(_nHandle,'   <NumberFormat ss:Format="[ENG][$-409]d\-mmm;@"/>' + c_ent)
fWrite(_nHandle,'  </Style>' + c_ent)
fWrite(_nHandle,'  <Style ss:ID="s27" ss:Parent="s16">' + c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>' + c_ent)
fWrite(_nHandle,'   <Borders/>' + c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="9" ss:Color="#000000"' + c_ent)
fWrite(_nHandle,'    ss:Bold="1"/>' + c_ent)
fWrite(_nHandle,'   <Interior/>' + c_ent)
fWrite(_nHandle,'   <NumberFormat ss:Format="[ENG][$-409]d\-mmm;@"/>' + c_ent)
fWrite(_nHandle,'  </Style>' + c_ent)
fWrite(_nHandle,'  <Style ss:ID="s28" ss:Parent="s16">' + c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="9" ss:Color="#000000"/>' + c_ent)
fWrite(_nHandle,'   <NumberFormat ss:Format="[ENG][$-409]d\-mmm;@"/>' + c_ent)
fWrite(_nHandle,'  </Style>' + c_ent)
fWrite(_nHandle,'  <Style ss:ID="s52" ss:Parent="s16">' + c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Right" ss:Vertical="Bottom"/>' + c_ent)
fWrite(_nHandle,'   <Borders>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Bottom" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Left" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Right" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Top" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'   </Borders>' + c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="9" ss:Bold="1"/>' + c_ent)
fWrite(_nHandle,'   <Interior/>' + c_ent)
fWrite(_nHandle,'   <NumberFormat ss:Format="#,##0.00"/>' + c_ent)
fWrite(_nHandle,'  </Style>' + c_ent)
fWrite(_nHandle,'  <Style ss:ID="s121" ss:Parent="s17">' + c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>' + c_ent)
fWrite(_nHandle,'   <Borders>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Bottom" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Left" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Right" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Top" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'   </Borders>' + c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="9" ss:Color="#000000"' + c_ent)
fWrite(_nHandle,'    ss:Bold="1"/>' + c_ent)
fWrite(_nHandle,'   <NumberFormat ss:Format="[ENG][$-409]d\-mmm;@"/>' + c_ent)
fWrite(_nHandle,'  </Style>' + c_ent)

fWrite(_nHandle,'  <Style ss:ID="s122" ss:Parent="s17">' + c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>' + c_ent)
fWrite(_nHandle,'   <Borders>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Bottom" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Left" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Right" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Top" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'   </Borders>' + c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="9" ss:Color="#000000"' + c_ent)
fWrite(_nHandle,'    ss:Bold="1"/>' + c_ent)
fWrite(_nHandle,'   <NumberFormat ss:Format="[ENG][$-409]d\-mmm;@"/>' + c_ent)
fWrite(_nHandle,'  </Style>' + c_ent)

fWrite(_nHandle,'  <Style ss:ID="s242" ss:Parent="s16">' + c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>' + c_ent)
fWrite(_nHandle,'   <Borders>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Bottom" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Left" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Right" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'    <Border ss:Position="Top" ss:LineStyle="Dot" ss:Weight="1"/>' + c_ent)
fWrite(_nHandle,'   </Borders>' + c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="9" ss:Color="#000000"' + c_ent)
fWrite(_nHandle,'    ss:Bold="1"/>' + c_ent)
fWrite(_nHandle,'   <Interior/>' + c_ent)
fWrite(_nHandle,'   <NumberFormat ss:Format="[ENG][$-409]d\-mmm;@"/>' + c_ent)
fWrite(_nHandle,'  </Style>' + c_ent)

fWrite(_nHandle,' </Styles>' + c_ent)

fWrite(_nHandle,' <Worksheet ss:Name="Movimento Anual ">' + c_ent)
fWrite(_nHandle,'  <Table ss:ExpandedColumnCount="30" ss:ExpandedRowCount="'+cValToChar(Len(aLinha)+3)+'" x:FullColumns="1"' + c_ent)
fWrite(_nHandle,'   x:FullRows="1" ss:StyleID="s19" ss:DefaultRowHeight="12">' + c_ent)
fWrite(_nHandle,'   <Column ss:StyleID="s19" ss:AutoFitWidth="0" ss:Width="16.5"/>' + c_ent)
fWrite(_nHandle,'   <Column ss:StyleID="s19" ss:AutoFitWidth="0" ss:Width="87"/>' + c_ent)
fWrite(_nHandle,'   <Column ss:StyleID="s20" ss:AutoFitWidth="0" ss:Width="214.5"/>' + c_ent)
fWrite(_nHandle,'   <Column ss:StyleID="s21" ss:AutoFitWidth="0" ss:Width="90.5" ss:Span="23"/>' + c_ent)
fWrite(_nHandle,'   <Column ss:Index="28" ss:StyleID="s22" ss:Width="25.5"/>' + c_ent)
fWrite(_nHandle,'   <Column ss:StyleID="s21" ss:AutoFitWidth="0" ss:Width="98.75" ss:Span="1"/>' + c_ent)

fWrite(_nHandle,'   <Row ss:AutoFitHeight="0" ss:Height="27.75" ss:StyleID="s28">' + c_ent)
fWrite(_nHandle,'    <Cell ss:Index="2" ss:StyleID="s121"/>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s24"/>' + c_ent)
fWrite(_nHandle,'    <Cell ss:MergeAcross="1" ss:StyleID="m300752156"><Data ss:Type="String">Janeiro</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:MergeAcross="1" ss:StyleID="m300752176"><Data ss:Type="String">Fevereiro</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:MergeAcross="1" ss:StyleID="m300752196"><Data ss:Type="String">Marco</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:MergeAcross="1" ss:StyleID="m300752216"><Data ss:Type="String">Abril</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:MergeAcross="1" ss:StyleID="m300752236"><Data ss:Type="String">Maio</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:MergeAcross="1" ss:StyleID="m300752256"><Data ss:Type="String">Junho</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:MergeAcross="1" ss:StyleID="m300752276"><Data ss:Type="String">Julho</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:MergeAcross="1" ss:StyleID="m300752296"><Data ss:Type="String">Agosto</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:MergeAcross="1" ss:StyleID="m300752316"><Data ss:Type="String">Setembro</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:MergeAcross="1" ss:StyleID="m300752760"><Data ss:Type="String">Outubro</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:MergeAcross="1" ss:StyleID="m300752780"><Data ss:Type="String">Novembro</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:MergeAcross="1" ss:StyleID="m300752800"><Data ss:Type="String">Dezembro</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s27"/>' + c_ent)
fWrite(_nHandle,'    <Cell ss:MergeAcross="1" ss:StyleID="m300752136"><Data ss:Type="String">Total Geral</Data></Cell>' + c_ent)
fWrite(_nHandle,'   </Row>' + c_ent)

fWrite(_nHandle,'   <Row ss:AutoFitHeight="0" ss:Height="27.75" ss:StyleID="s28">' + c_ent)
fWrite(_nHandle,'    <Cell ss:Index="2" ss:StyleID="s121"><Data ss:Type="String" x:Ticked="1">Codigo</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s24"><Data ss:Type="String">Descricao</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s26"><Data ss:Type="String">'+cDesc1+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s26"><Data ss:Type="String">'+cDesc2+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s26"><Data ss:Type="String">'+cDesc1+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s26"><Data ss:Type="String">'+cDesc2+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s26"><Data ss:Type="String">'+cDesc1+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s26"><Data ss:Type="String">'+cDesc2+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s26"><Data ss:Type="String">'+cDesc1+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s26"><Data ss:Type="String">'+cDesc2+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s26"><Data ss:Type="String">'+cDesc1+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s26"><Data ss:Type="String">'+cDesc2+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s26"><Data ss:Type="String">'+cDesc1+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s26"><Data ss:Type="String">'+cDesc2+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s26"><Data ss:Type="String">'+cDesc1+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s26"><Data ss:Type="String">'+cDesc2+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s26"><Data ss:Type="String">'+cDesc1+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s26"><Data ss:Type="String">'+cDesc2+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s26"><Data ss:Type="String">'+cDesc1+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s26"><Data ss:Type="String">'+cDesc2+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s26"><Data ss:Type="String">'+cDesc1+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s26"><Data ss:Type="String">'+cDesc2+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s26"><Data ss:Type="String">'+cDesc1+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s26"><Data ss:Type="String">'+cDesc2+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s26"><Data ss:Type="String">'+cDesc1+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s26"><Data ss:Type="String">'+cDesc2+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s27"/>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s26"><Data ss:Type="String">'+cDesc1+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s26"><Data ss:Type="String">'+cDesc2+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'   </Row>' + c_ent)

ProcRegua(Len(aLinha))

For nCont := 1 to Len(aLinha)
	IncProc("Gravando a Linha "+cValToChar(nCont)+" / "+cValToChar(Len(aLinha)))
	if Alltrim(aLinha[nCont,3]) <> ""
		fWrite(_nHandle,'   <Row>' + c_ent)
		fWrite(_nHandle,'    <Cell ss:Index="2" ss:StyleID="s122"><Data ss:Type="String">'+Space(Val(aLinha[nCont,1])*3)+aLinha[nCont,3]+'</Data></Cell>' + c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s242"><Data ss:Type="String">'+Space(Val(aLinha[nCont,1])*3)+aLinha[nCont,4]+'</Data></Cell>' + c_ent)
		For nCont2 := 5 to 28
			fWrite(_nHandle,'    <Cell ss:StyleID="s52"><Data ss:Type="Number">'+cValToChar(aLinha[nCont,nCont2])+'</Data></Cell>' + c_ent)
		Next
		fWrite(_nHandle,'    <Cell ss:StyleID="s52"/>' + c_ent)
		For nCont2 := 29 to 30
			fWrite(_nHandle,'    <Cell ss:StyleID="s52"><Data ss:Type="Number">'+cValToChar(aLinha[nCont,nCont2])+'</Data></Cell>' + c_ent)
		Next

		fWrite(_nHandle,'   </Row>' + c_ent)
	endif
Next

fWrite(_nHandle,'   <Row>' + c_ent)
fWrite(_nHandle,'    <Cell ss:Index="2" ss:StyleID="s121"><Data ss:Type="String">Total</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s24"/>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s52"><Data ss:Type="Number">'+cValToChar(nTotal01)+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s52"><Data ss:Type="Number">'+cValToChar(nTotal02)+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s52"><Data ss:Type="Number">'+cValToChar(nTotal03)+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s52"><Data ss:Type="Number">'+cValToChar(nTotal04)+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s52"><Data ss:Type="Number">'+cValToChar(nTotal05)+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s52"><Data ss:Type="Number">'+cValToChar(nTotal06)+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s52"><Data ss:Type="Number">'+cValToChar(nTotal07)+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s52"><Data ss:Type="Number">'+cValToChar(nTotal08)+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s52"><Data ss:Type="Number">'+cValToChar(nTotal09)+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s52"><Data ss:Type="Number">'+cValToChar(nTotal10)+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s52"><Data ss:Type="Number">'+cValToChar(nTotal11)+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s52"><Data ss:Type="Number">'+cValToChar(nTotal12)+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s52"><Data ss:Type="Number">'+cValToChar(nTotal13)+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s52"><Data ss:Type="Number">'+cValToChar(nTotal14)+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s52"><Data ss:Type="Number">'+cValToChar(nTotal15)+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s52"><Data ss:Type="Number">'+cValToChar(nTotal16)+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s52"><Data ss:Type="Number">'+cValToChar(nTotal17)+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s52"><Data ss:Type="Number">'+cValToChar(nTotal18)+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s52"><Data ss:Type="Number">'+cValToChar(nTotal19)+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s52"><Data ss:Type="Number">'+cValToChar(nTotal20)+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s52"><Data ss:Type="Number">'+cValToChar(nTotal21)+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s52"><Data ss:Type="Number">'+cValToChar(nTotal22)+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s52"><Data ss:Type="Number">'+cValToChar(nTotal23)+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s52"><Data ss:Type="Number">'+cValToChar(nTotal24)+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s52"/>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s52"><Data ss:Type="Number">'+cValToChar(nTotal25)+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s52"><Data ss:Type="Number">'+cValToChar(nTotal26)+'</Data></Cell>' + c_ent)
fWrite(_nHandle,'   </Row>' + c_ent)

fWrite(_nHandle,'  </Table>' + c_ent)

fWrite(_nHandle,'  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">' + c_ent)
fWrite(_nHandle,'   <PageSetup>' + c_ent)
fWrite(_nHandle,'    <Layout x:Orientation="Landscape"/>' + c_ent)
fWrite(_nHandle,'    <Header x:Margin="0.31496062992125984"/>' + c_ent)
fWrite(_nHandle,'    <Footer x:Margin="0.31496062992125984"/>' + c_ent)
fWrite(_nHandle,'    <PageMargins x:Bottom="0.78740157480314965" x:Left="0.11811023622047245"' + c_ent)
fWrite(_nHandle,'     x:Right="0.11811023622047245" x:Top="0.78740157480314965"/>' + c_ent)
fWrite(_nHandle,'   </PageSetup>' + c_ent)
fWrite(_nHandle,'   <NoSummaryRowsBelowDetail/>' + c_ent)
fWrite(_nHandle,'   <FitToPage/>' + c_ent)
fWrite(_nHandle,'   <Print>' + c_ent)
fWrite(_nHandle,'    <ValidPrinterInfo/>' + c_ent)
fWrite(_nHandle,'    <Scale>54</Scale>' + c_ent)
fWrite(_nHandle,'    <HorizontalResolution>300</HorizontalResolution>' + c_ent)
fWrite(_nHandle,'    <VerticalResolution>300</VerticalResolution>' + c_ent)
fWrite(_nHandle,'   </Print>' + c_ent)
fWrite(_nHandle,'   <Zoom>90</Zoom>' + c_ent)
fWrite(_nHandle,'   <Selected/>' + c_ent)
fWrite(_nHandle,'   <DoNotDisplayGridlines/>' + c_ent)
fWrite(_nHandle,'   <FreezePanes/>' + c_ent)
fWrite(_nHandle,'   <FrozenNoSplit/>' + c_ent)
fWrite(_nHandle,'   <SplitHorizontal>1</SplitHorizontal>' + c_ent)
fWrite(_nHandle,'   <TopRowBottomPane>1</TopRowBottomPane>' + c_ent)
fWrite(_nHandle,'   <SplitVertical>3</SplitVertical>' + c_ent)
fWrite(_nHandle,'   <LeftColumnRightPane>3</LeftColumnRightPane>' + c_ent)
fWrite(_nHandle,'   <ActivePane>0</ActivePane>' + c_ent)
fWrite(_nHandle,'   <Panes>' + c_ent)
fWrite(_nHandle,'    <Pane>' + c_ent)
fWrite(_nHandle,'     <Number>3</Number>' + c_ent)
fWrite(_nHandle,'     <ActiveRow>202</ActiveRow>' + c_ent)
fWrite(_nHandle,'     <ActiveCol>5</ActiveCol>' + c_ent)
fWrite(_nHandle,'    </Pane>' + c_ent)
fWrite(_nHandle,'    <Pane>' + c_ent)
fWrite(_nHandle,'     <Number>1</Number>' + c_ent)
fWrite(_nHandle,'     <ActiveRow>202</ActiveRow>' + c_ent)
fWrite(_nHandle,'     <ActiveCol>5</ActiveCol>' + c_ent)
fWrite(_nHandle,'    </Pane>' + c_ent)
fWrite(_nHandle,'    <Pane>' + c_ent)
fWrite(_nHandle,'     <Number>2</Number>' + c_ent)
fWrite(_nHandle,'     <ActiveRow>202</ActiveRow>' + c_ent)
fWrite(_nHandle,'     <ActiveCol>5</ActiveCol>' + c_ent)
fWrite(_nHandle,'    </Pane>' + c_ent)
fWrite(_nHandle,'    <Pane>' + c_ent)
fWrite(_nHandle,'     <Number>0</Number>' + c_ent)
fWrite(_nHandle,'     <ActiveRow>0</ActiveRow>' + c_ent)
fWrite(_nHandle,'     <ActiveCol>0</ActiveCol>' + c_ent)
fWrite(_nHandle,'    </Pane>' + c_ent)
fWrite(_nHandle,'   </Panes>' + c_ent)
fWrite(_nHandle,'   <ProtectObjects>False</ProtectObjects>' + c_ent)
fWrite(_nHandle,'   <ProtectScenarios>False</ProtectScenarios>' + c_ent)
fWrite(_nHandle,'  </WorksheetOptions>' + c_ent)
fWrite(_nHandle,' </Worksheet>' + c_ent)
fWrite(_nHandle,'</Workbook>' + c_ent)

dbSelectArea(cAlias)
(cAlias)->(dbCloseArea())

//****************|
//Fechando arquivo|
//****************|
fClose(_nHandle)

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³Abre a Planilha Excel³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
ShellExecute("Open",_cArquivo," ",_cPath,3)

Return
