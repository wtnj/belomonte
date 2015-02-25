#INCLUDE "Protheus.ch"
#INCLUDE "TopConn.ch"
#INCLUDE "rwmake.ch" 

#DEFINE c_ent CHR(13)+CHR(10)

#DEFINE _OPC_cGETFILE ( GETF_RETDIRECTORY + GETF_ONLYSERVER )


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³OIFR004   ºAutor  ³Leandro Ribeiro     º Data ³  13/06/14   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Programa para gerar planilha de conta com movimentações    º±±
±±º          ³ na tabela AKD de acordo com a parametrização do usuario.   º±±
±±º          ³ Planilha Base de Extração								  º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP                                                        º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

User Function OIFR004()  

Local _aArea := GetArea() 
Local cPerg  := Padr("OIFR004",Len(SX1->X1_GRUPO))      	             
                                                                                      //AL2A                                                                                                     ENG
//PutSx1(cPerg,"01","Tipo de Saldo"         , "", "", "mv_ch1" , "C", 08,0 ,0, "C", "", ""      ,"", "", "MV_PAR01", "PG" , "" , "" , "", "RE"      ,"","","EM","","","CT","","","","","",""/*,"","",""*///)     
PutSx1(cPerg,"01","Tipo de Saldo"             , "", "", "mv_ch1"  , "C", TAMSX3("AKD_TPSALD")[1]   ,0 ,0, "G", "", "AL2A"    ,"", "", "MV_PAR01" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"")
PutSx1(cPerg,"02","Item Contabil de?"         , "", "", "mv_ch2"  , "C", TAMSX3("AKD_ITCTB")[1]    ,0 ,0, "G", "", "CTD"     ,"", "", "MV_PAR02" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"")
PutSx1(cPerg,"03","Item Contabil Ate?"        , "", "", "mv_ch3"  , "C", TAMSX3("AKD_ITCTB")[1]    ,0 ,0, "G", "", "CTD"     ,"", "", "MV_PAR03" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"")
PutSx1(cPerg,"04","Centro de Custo de?"       , "", "", "mv_ch4"  , "C", TAMSX3("AKD_CC")[1]       ,0, 0, "G", "", "CTT"     ,"", "", "MV_PAR04" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"") 
PutSx1(cPerg,"05","Centro de Custo ate?"      , "", "", "mv_ch5"  , "C", TAMSX3("AKD_CC")[1]       ,0, 0, "G", "", "CTT"     ,"", "", "MV_PAR05" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"") 
PutSx1(cPerg,"06","Oper. Orçamentaria de?" 	  , "", "", "mv_ch6"  , "C", TAMSX3("AKF_CODIGO")[1]   ,0, 0, "G", "", "AKF"     ,"", "", "MV_PAR06" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"")
PutSx1(cPerg,"07","Oper. Orçamentaria Ate?"	  , "", "", "mv_ch7"  , "C", TAMSX3("AKF_CODIGO")[1]   ,0, 0, "G", "", "AKF"     ,"", "", "MV_PAR07" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"")
PutSx1(cPerg,"08","Data de"               	  , "", "", "mv_ch8"  , "D", TAMSX3("AKD_DATA")[1]     ,0, 0, "G", "", ""        ,"", "", "MV_PAR08" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"")
PutSx1(cPerg,"09","Data Ate"              	  , "", "", "mv_ch9"  , "D", TAMSX3("AKD_DATA")[1]     ,0, 0, "G", "", ""        ,"", "", "MV_PAR09" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"")
PutSx1(cPerg,"10","Classe Orçamentaria?"      , "", "", "mv_ch10"  , "C", TAMSX3("AK6_CODIGO")[1]   ,0, 0, "G", "", "AK6"     ,"", "", "MV_PAR10" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"") 
PutSx1(cPerg,"11","Conta Orçamentaria de?"    , "", "", "mv_ch11" , "C", TAMSX3("AK5_CODIGO")[1]   ,0, 0, "G", "", "AK5"     ,"", "", "MV_PAR11" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"") 
PutSx1(cPerg,"12","Conta Orçamentaria ate?"   , "", "", "mv_ch12" , "C", TAMSX3("AK5_CODIGO")[1]   ,0, 0, "G", "", "AK5"     ,"", "", "MV_PAR12" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"") 
//PutSx1(cPerg,"12","Conta Contábil de?"        , "", "", "mv_ch12" , "C", TAMSX3("CT1_CONTA")[1]    ,0, 0, "G", "", "CT1"     ,"", "", "MV_PAR12" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"") 
//PutSx1(cPerg,"13","Conta Contábil ate?"       , "", "", "mv_ch13" , "C", TAMSX3("CT1_CONTA")[1]    ,0, 0, "G", "", "CT1"     ,"", "", "MV_PAR13" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"")                                                                             
PutSx1(cPerg,"13","Fornecedor de?"            , "", "", "mv_ch13" , "C", TAMSX3("A2_COD")[1]       ,0, 0, "G", "", "SA2"     ,"", "", "MV_PAR13" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"") 
PutSx1(cPerg,"14","Fornecedor ate?"           , "", "", "mv_ch14" , "C", TAMSX3("A2_COD")[1]       ,0, 0, "G", "", "SA2"     ,"", "", "MV_PAR14" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"") 
                             
If Pergunte(cPerg,.T.)
	
	Processa({|| GOIFR004() },"Aguarde...")
	
EndIf

RestArea(_aArea)
                                               
Return()


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³GOIFR004   ºAutor  ³Leandro Ribeiro    º Data ³  13/06/14   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Função para gerar a planilha de excel                      º±±
±±º          ³                                                            º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP                                                        º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function GOIFR004()

Local _cOIFR004 := GetArea()
//Local _cFile    := "OIFR004" + DToS( Date() ) + ".xml"                      
Local _cFile    := "OIFR004.xml"                      
Local _cPath    := GetTempPath() //Pega a Pasta de Arquivo temporário                 

Local _cQuery0  := "" 
Local _cTrab0  	:= GetNextAlias() 
Local _cArray0  := {} //Movimentos na Tabela AKD
Local _cFornece	:= RETFORNECE(MV_PAR13,MV_PAR14) // Retornas os Fornecedores
//Local _cContCTB	:= RETCONTCTB(MV_PAR12,MV_PAR13) // Retornas as Contas Contabeis 
Local _cArray3  := {}  
Local _cArray5  := {} 
Local _cItSint	:= "" 
Local _cTotal   := 0
Local _cConta   := 0

Private cPath    := "" 
Private _nHandle 
                      
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³Após as validações começar a montar o relatório, gravando em XML³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³Monta o Nome do Arquivo³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
_cArquivo := Alltrim( _cPath ) + Iif( Right( Alltrim( _cPath ), 1 ) == "\", "", "\" ) + Alltrim( _cFile )
If File( _cArquivo )
	FErase( _cArquivo )
End If

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³Cria o Arquivo Fisico³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
_nHandle := MsFCreate( _cArquivo )     

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ//
//"Tipo de Saldo"             MV_PAR01  //
//"Item Contabil de?"         MV_PAR02  //
//"Item Contabil Ate?"        MV_PAR03  //
//"Centro de Custo de?"       MV_PAR04  //
//"Centro de Custo ate?"      MV_PAR05  //
//"Oper. Orçamentaria de?"    MV_PAR06  //
//"Oper. Orçamentaria ate?"   MV_PAR07  //
//"Data de"                   MV_PAR08  //
//"Data Ate"                  MV_PAR09  //
//"Classe Orçamentaria?"      MV_PAR10  //
//"Conta Orçamentaria de?"    MV_PAR11 //
//"Conta Orçamentaria ate?"   MV_PAR12 //
//"Fornecedor de?"            MV_PAR13 //
//"Fornecedor ate?"           MV_PAR14 //
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ//

_cQuery0  := " SELECT DISTINCT AKD_TPSALD, AKD_CLASSE, AKD_DATA, AKD_ITCTB, AKD_CC, AKD_OPER, AKD_CHAVE, AKD_CO, AKD_TIPO, AKD_VALOR1 "+ c_ent
_cQuery0  += " FROM "+RETSQLNAME("AKD")+" AKD "+ c_ent
_cQuery0  += " WHERE "+ c_ent
_cQuery0  += " AKD_FILIAL = '"+xFilial("AKD")+"' "+ c_ent
_cQuery0  += " AND AKD_TPSALD = '"+MV_PAR01+"' "+ c_ent
_cQuery0  += " AND AKD_ITCTB BETWEEN '"+MV_PAR02+"' AND '"+MV_PAR03+"' "+ c_ent
_cQuery0  += " AND AKD_CLASSE = '"+MV_PAR10+"' "+ c_ent
_cQuery0  += " AND AKD_CC BETWEEN '"+MV_PAR04+"' AND '"+MV_PAR05+"' "+ c_ent
_cQuery0  += " AND AKD_CO BETWEEN '"+MV_PAR11+"' AND '"+MV_PAR12+"' "+ c_ent  
_cQuery0  += " AND AKD_DATA BETWEEN '"+DTOS(MV_PAR08)+"' AND '"+DTOS(MV_PAR09)+"' "+ c_ent
//_cQuery0  += " AND AKD_OPER = '"+AllTrim(MV_PAR06)+"' "+ c_ent 
_cQuery0  += " AND AKD_STATUS = '1' "+ c_ent 
//_cQuery0  += " AND AKD_CHAVE LIKE '%MATA100%' "+ c_ent 
_cQuery0  += " AND AKD_OPER BETWEEN '"+MV_PAR06+"' AND '"+MV_PAR07+"' "+ c_ent
_cQuery0  += " AND AKD.D_E_L_E_T_ = ' ' "+ c_ent
_cQuery0  += " ORDER BY AKD_DATA "+ c_ent
_cQuery0 := ChangeQuery(_cQuery0) 

If Select(_cTrab0) > 0
	dbSelectArea(_cTrab0)
	(_cTrab0)->(dbCloseArea())
EndIf

dbUseArea(.T.,"TOPCONN",TcGenQry(,,_cQuery0),_cTrab0,.T.,.T.)       

dbSelectArea(_cTrab0)

If(!Eof())
  While !(_cTrab0)->(Eof())  
		Aadd(_cArray0,{(_cTrab0)->AKD_TPSALD,(_cTrab0)->AKD_CLASSE,(_cTrab0)->AKD_DATA,(_cTrab0)->AKD_ITCTB,(_cTrab0)->AKD_CC,(_cTrab0)->AKD_OPER,(_cTrab0)->AKD_CHAVE,(_cTrab0)->AKD_CO,(_cTrab0)->AKD_TIPO,(_cTrab0)->AKD_VALOR1})
  	DbSkip()
  EndDo 
EndIf  

(_cTrab0)->(dbCloseArea())    


If(EMPTY(_cArray0))
   Aviso("Aviso","Nenhum dado encontrado com os parametros passados!",{"OK"})   
   	U_OIFR004()
   Return()
EndIf                    

_cTotal := Len(_cArray0)
ProcRegua(_cTotal)                                                                                                              

For _cc := 1 to Len(_cArray0)  

	_cConta := _cConta + 1 
		
	IncProc("Verificando Registros - Linha: "+Alltrim(str(_cConta)) +" de "+ Alltrim(str(_cTotal)) +" ...")   
	
	_cArray5 := RETCHAVE(_cArray0[_cc][7],_cFornece)  	         
     
 	If(!Empty(_cArray5))
		For _xx := 1 to Len(_cArray5)				
		   	Aadd(_cArray3,{_cArray0[_cc][1],;
			               _cArray0[_cc][2],;
			               ReturnMes(STOD(_cArray0[_cc][3])),;
		   	               _cArray0[_cc][3],; 	               
		   	               Iif(!EMPTY(_cArray0[_cc][4]),TiraGraf(Alltrim(POSICIONE("CTD",1,xFilial("CTD")+PADR(AllTrim(_cArray0[_cc][4]),TAMSX3("CTD_ITEM")[1]),"CTD_DESC01"))),""),;
		   	               Iif(!EMPTY(_cArray0[_cc][4]),TiraGraf(RETITSINT(_cArray0[_cc][4])),""),;
						   Iif(!EMPTY(_cArray0[_cc][5]),TiraGraf(Alltrim(POSICIONE("CTT",1,xFilial("CTT")+PADR(AllTrim(_cArray0[_cc][5]),TAMSX3("CTT_CUSTO")[1]),"CTT_DESC01"))),""),;
						   Iif(!EMPTY(_cArray0[_cc][6]),TiraGraf(Alltrim(POSICIONE("AKF",1,xFilial("AKF")+PADR(AllTrim(_cArray0[_cc][6]),TAMSX3("AKF_CODIGO")[1]),"AKF_DESCRI"))),""),;
						   _cArray5[_xx][1],_cArray5[_xx][2],_cArray5[_xx][3],_cArray5[_xx][4],_cArray5[_xx][5],_cArray5[_xx][6],;
		   				   _cArray5[_xx][7],_cArray0[_cc][8],;
						   Iif(!EMPTY(_cArray0[_cc][8]),TiraGraf(Alltrim(POSICIONE("AK5",1,xFilial("AK5")+PADR(AllTrim(_cArray0[_cc][8]),TAMSX3("AK5_CODIGO")[1]),"AK5_DESCRI"))),""),;
			               _cArray5[_xx][8],_cArray5[_xx][9],_cArray5[_xx][10],_cArray5[_xx][11],_cArray5[_xx][12],_cArray5[_xx][13],;
		   				   _cArray5[_xx][14],_cArray5[_xx][15],_cArray5[_xx][16],_cArray5[_xx][17],_cArray5[_xx][18],_cArray0[_cc][9],_cArray0[_cc][10]})  
	   	
	   	Next _xx
	 Else 
	 	Aadd(_cArray3,{_cArray0[_cc][1],;
			               _cArray0[_cc][2],;
			               ReturnMes(STOD(_cArray0[_cc][3])),;
		   	               _cArray0[_cc][3],; 	               
		   	               Iif(!EMPTY(_cArray0[_cc][4]),TiraGraf(Alltrim(POSICIONE("CTD",1,xFilial("CTD")+PADR(AllTrim(_cArray0[_cc][4]),TAMSX3("CTD_ITEM")[1]),"CTD_DESC01"))),""),;
		   	               Iif(!EMPTY(_cArray0[_cc][4]),TiraGraf(RETITSINT(_cArray0[_cc][4])),""),;
						   Iif(!EMPTY(_cArray0[_cc][5]),TiraGraf(Alltrim(POSICIONE("CTT",1,xFilial("CTT")+PADR(AllTrim(_cArray0[_cc][5]),TAMSX3("CTT_CUSTO")[1]),"CTT_DESC01"))),""),;
						   Iif(!EMPTY(_cArray0[_cc][6]),TiraGraf(Alltrim(POSICIONE("AKF",1,xFilial("AKF")+PADR(AllTrim(_cArray0[_cc][6]),TAMSX3("AKF_CODIGO")[1]),"AKF_DESCRI"))),""),;
						   "","","","","","",;
		   				   "",_cArray0[_cc][8],;
						   Iif(!EMPTY(_cArray0[_cc][8]),TiraGraf(Alltrim(POSICIONE("AK5",1,xFilial("AK5")+PADR(AllTrim(_cArray0[_cc][8]),TAMSX3("AK5_CODIGO")[1]),"AK5_DESCRI"))),""),;
			               "","","","","","","","","","","",_cArray0[_cc][9],_cArray0[_cc][10]})  
	   	EndIf  
   				   
   	_cArray5 := {}

Next _cc  

_cConta := 0

If(EMPTY(_cArray3))
   Aviso("Aviso","Nenhum dado encontrado com os parametros passados!",{"OK"})   
   	U_OIFR004()
   Return()
EndIf   
            
fWrite(_nHandle,'	<?xml version="1.0"?>   	'+ c_ent)
fWrite(_nHandle,'	<?mso-application progid="Excel.Sheet"?>	'+ c_ent)
fWrite(_nHandle,'	<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"	'+ c_ent)
fWrite(_nHandle,'	 xmlns:o="urn:schemas-microsoft-com:office:office"	'+ c_ent)
fWrite(_nHandle,'	 xmlns:x="urn:schemas-microsoft-com:office:excel"	'+ c_ent)
fWrite(_nHandle,'	 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"	'+ c_ent)
fWrite(_nHandle,'	 xmlns:html="http://www.w3.org/TR/REC-html40">	'+ c_ent)
fWrite(_nHandle,'	 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">	'+ c_ent)
fWrite(_nHandle,'	  <Author>PROTHEUS 11</Author>	'+ c_ent)
fWrite(_nHandle,'	  <LastAuthor>PROTHEUS 11</LastAuthor>	'+ c_ent)
fWrite(_nHandle,'	  <Created>2013-05-23T13:58:13Z</Created>	'+ c_ent)
fWrite(_nHandle,'	  <LastSaved>2014-06-21T19:43:26Z</LastSaved>	'+ c_ent)
fWrite(_nHandle,'	  <Company>Microsoft</Company>	'+ c_ent)
fWrite(_nHandle,'	  <Version>12.00</Version>	'+ c_ent)
fWrite(_nHandle,'	 </DocumentProperties>	'+ c_ent)
fWrite(_nHandle,'	 <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">	'+ c_ent)
fWrite(_nHandle,'	  <SupBook>	'+ c_ent)
fWrite(_nHandle,'	   <Path>extracao</Path>	'+ c_ent)
fWrite(_nHandle,'	   <SheetName>extracao</SheetName>	'+ c_ent)
fWrite(_nHandle,'	  </SupBook>	'+ c_ent)
fWrite(_nHandle,'	  <WindowHeight>7050</WindowHeight>	'+ c_ent)
fWrite(_nHandle,'	  <WindowWidth>11955</WindowWidth>	'+ c_ent)
fWrite(_nHandle,'	  <WindowTopX>240</WindowTopX>	'+ c_ent)
fWrite(_nHandle,'	  <WindowTopY>360</WindowTopY>	'+ c_ent)
fWrite(_nHandle,'	  <TabRatio>921</TabRatio>	'+ c_ent)
fWrite(_nHandle,'	  <ProtectStructure>False</ProtectStructure>	'+ c_ent)
fWrite(_nHandle,'	  <ProtectWindows>False</ProtectWindows>	'+ c_ent)
fWrite(_nHandle,'	 </ExcelWorkbook>	'+ c_ent)
fWrite(_nHandle,'	 <Styles>	'+ c_ent)
fWrite(_nHandle,'	  <Style ss:ID="Default" ss:Name="Normal">	'+ c_ent)
fWrite(_nHandle,'	   <Alignment ss:Vertical="Bottom"/>	'+ c_ent)
fWrite(_nHandle,'	   <Borders/>	'+ c_ent)
fWrite(_nHandle,'	   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>	'+ c_ent)
fWrite(_nHandle,'	   <Interior/>	'+ c_ent)
fWrite(_nHandle,'	   <NumberFormat/>	'+ c_ent)
fWrite(_nHandle,'	   <Protection/>	'+ c_ent)
fWrite(_nHandle,'	  </Style>	'+ c_ent)
fWrite(_nHandle,'	  <Style ss:ID="s62" ss:Name="Moeda 3">	'+ c_ent)
fWrite(_nHandle,'	   <NumberFormat	'+ c_ent)
fWrite(_nHandle,'	    ss:Format="_(&quot;R$ &quot;* #,##0.00_);_(&quot;R$ &quot;* \(#,##0.00\);_(&quot;R$ &quot;* &quot;-&quot;??_);_(@_)"/>	'+ c_ent)
fWrite(_nHandle,'	  </Style>	'+ c_ent)
fWrite(_nHandle,'	  <Style ss:ID="s63">	'+ c_ent)
fWrite(_nHandle,'	   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>	'+ c_ent)
fWrite(_nHandle,'	   <Interior/>	'+ c_ent)
fWrite(_nHandle,'	  </Style>	'+ c_ent)
fWrite(_nHandle,'	  <Style ss:ID="s64">	'+ c_ent)
fWrite(_nHandle,'	   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>	'+ c_ent)
fWrite(_nHandle,'	   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>	'+ c_ent)
fWrite(_nHandle,'	   <Interior/>	'+ c_ent)
fWrite(_nHandle,'	  </Style>	'+ c_ent)
fWrite(_nHandle,'	  <Style ss:ID="s65">	'+ c_ent)
fWrite(_nHandle,'	   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom" ss:Indent="8"/>	'+ c_ent)
fWrite(_nHandle,'	   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Color="#000000" ss:Bold="1"/>	'+ c_ent)
fWrite(_nHandle,'	  </Style>	'+ c_ent)
fWrite(_nHandle,'	  <Style ss:ID="s66">	'+ c_ent)
fWrite(_nHandle,'	   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"	'+ c_ent)
fWrite(_nHandle,'	    ss:Bold="1"/>	'+ c_ent)
fWrite(_nHandle,'	  </Style>	'+ c_ent)
fWrite(_nHandle,'	  <Style ss:ID="s67">	'+ c_ent)
fWrite(_nHandle,'	   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"	'+ c_ent)
fWrite(_nHandle,'	    ss:Bold="1"/>	'+ c_ent)
fWrite(_nHandle,'	   <NumberFormat ss:Format="_-* #,##0.00_-;\-* #,##0.00_-;_-* &quot;-&quot;??_-;_-@_-"/>	'+ c_ent)
fWrite(_nHandle,'	  </Style>	'+ c_ent)
fWrite(_nHandle,'	  <Style ss:ID="s68">	'+ c_ent)
fWrite(_nHandle,'	   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>	'+ c_ent)
fWrite(_nHandle,'	   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>	'+ c_ent)
fWrite(_nHandle,'	  </Style>	'+ c_ent)
fWrite(_nHandle,'	  <Style ss:ID="s69">	'+ c_ent)
fWrite(_nHandle,'	   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>	'+ c_ent)
fWrite(_nHandle,'	   <NumberFormat	'+ c_ent)
fWrite(_nHandle,'	    ss:Format="_-&quot;R$&quot;\ * #,##0.00_-;\-&quot;R$&quot;\ * #,##0.00_-;_-&quot;R$&quot;\ * &quot;-&quot;??_-;_-@_-"/>	'+ c_ent)
fWrite(_nHandle,'	  </Style>	'+ c_ent)
fWrite(_nHandle,'	  <Style ss:ID="s70">	'+ c_ent)
fWrite(_nHandle,'	   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>	'+ c_ent)
fWrite(_nHandle,'	   <NumberFormat	'+ c_ent)
fWrite(_nHandle,'	    ss:Format="_-&quot;R$&quot;\ * #,##0_-;\-&quot;R$&quot;\ * #,##0_-;_-&quot;R$&quot;\ * &quot;-&quot;??_-;_-@_-"/>	'+ c_ent)
fWrite(_nHandle,'	  </Style>	'+ c_ent)
fWrite(_nHandle,'	  <Style ss:ID="s71">	'+ c_ent)
fWrite(_nHandle,'	   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#0066CC"/>	'+ c_ent)
fWrite(_nHandle,'	   <Interior/>	'+ c_ent)
fWrite(_nHandle,'	  </Style>	'+ c_ent)
fWrite(_nHandle,'	  <Style ss:ID="s73">	'+ c_ent)
fWrite(_nHandle,'	   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>	'+ c_ent)
fWrite(_nHandle,'	   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#0066CC"/>	'+ c_ent)
fWrite(_nHandle,'	  </Style>	'+ c_ent)
fWrite(_nHandle,'	  <Style ss:ID="s74">	'+ c_ent)
fWrite(_nHandle,'	   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>	'+ c_ent)
fWrite(_nHandle,'	  </Style>	'+ c_ent)
fWrite(_nHandle,'	  <Style ss:ID="s75">	'+ c_ent)
fWrite(_nHandle,'	   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#0066CC"/>	'+ c_ent)
fWrite(_nHandle,'	   <NumberFormat ss:Format="#,##0"/>	'+ c_ent)
fWrite(_nHandle,'	  </Style>	'+ c_ent)
fWrite(_nHandle,'	  <Style ss:ID="s76">	'+ c_ent)
fWrite(_nHandle,'	   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>	'+ c_ent)
fWrite(_nHandle,'	   <Borders>	'+ c_ent)
fWrite(_nHandle,'	    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>	'+ c_ent)
fWrite(_nHandle,'	    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>	'+ c_ent)
fWrite(_nHandle,'	    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>	'+ c_ent)
fWrite(_nHandle,'	    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>	'+ c_ent)
fWrite(_nHandle,'	   </Borders>	'+ c_ent)
fWrite(_nHandle,'	   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#FFFFFF"	'+ c_ent)
fWrite(_nHandle,'	    ss:Bold="1"/>	'+ c_ent)
fWrite(_nHandle,'	   <Interior ss:Color="#1D1B11" ss:Pattern="Solid"/>	'+ c_ent)
fWrite(_nHandle,'	  </Style>	'+ c_ent)
fWrite(_nHandle,'	  <Style ss:ID="s77">	'+ c_ent)
fWrite(_nHandle,'	   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>	'+ c_ent)
fWrite(_nHandle,'	   <Borders>	'+ c_ent)
fWrite(_nHandle,'	    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>	'+ c_ent)
fWrite(_nHandle,'	    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>	'+ c_ent)
fWrite(_nHandle,'	    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>	'+ c_ent)
fWrite(_nHandle,'	    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>	'+ c_ent)
fWrite(_nHandle,'	   </Borders>	'+ c_ent)
fWrite(_nHandle,'	   <Interior/>	'+ c_ent)
fWrite(_nHandle,'	  </Style>	'+ c_ent)
fWrite(_nHandle,'	  <Style ss:ID="s78">	'+ c_ent)
//fWrite(_nHandle,'	   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>	'+ c_ent)
fWrite(_nHandle,'	   <Borders>	'+ c_ent)
fWrite(_nHandle,'	    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>	'+ c_ent)
fWrite(_nHandle,'	    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>	'+ c_ent)
fWrite(_nHandle,'	    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>	'+ c_ent)
fWrite(_nHandle,'	    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>	'+ c_ent)
fWrite(_nHandle,'	   </Borders>	'+ c_ent)
fWrite(_nHandle,'	   <Interior/>	'+ c_ent)
fWrite(_nHandle,'	   <NumberFormat ss:Format="Short Date"/>	'+ c_ent)
fWrite(_nHandle,'	     </Style>	'+ c_ent)
fWrite(_nHandle,'	     <Style ss:ID="s79">	'+ c_ent)
fWrite(_nHandle,'	      <Borders>	'+ c_ent)
fWrite(_nHandle,'	       <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>	'+ c_ent)
fWrite(_nHandle,'	       <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>	'+ c_ent)
fWrite(_nHandle,'	       <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>	'+ c_ent)
fWrite(_nHandle,'	       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>	'+ c_ent)
fWrite(_nHandle,'	      </Borders>	'+ c_ent)
fWrite(_nHandle,'	      <Interior/>	'+ c_ent)
fWrite(_nHandle,'	     </Style>	'+ c_ent) 
fWrite(_nHandle,'  <Style ss:ID="s81" ss:Parent="s62">'+ c_ent) 
fWrite(_nHandle,'   <Borders>'+ c_ent) 
fWrite(_nHandle,'    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>'+ c_ent) 
fWrite(_nHandle,'    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>'+ c_ent) 
fWrite(_nHandle,'    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>'+ c_ent) 
fWrite(_nHandle,'    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>'+ c_ent) 
fWrite(_nHandle,'   </Borders>'+ c_ent) 
fWrite(_nHandle,'   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>'+ c_ent) 
fWrite(_nHandle,'   <Interior/>'+ c_ent) 
fWrite(_nHandle,'   <NumberFormat ss:Format="@"/>'+ c_ent) 
fWrite(_nHandle,'  </Style>'+ c_ent) 
fWrite(_nHandle,'	     <Style ss:ID="s82" ss:Parent="s62">	'+ c_ent)
fWrite(_nHandle,'	      <Borders>	'+ c_ent)
fWrite(_nHandle,'	       <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>	'+ c_ent)
fWrite(_nHandle,'	       <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>	'+ c_ent)
fWrite(_nHandle,'	       <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>	'+ c_ent)
fWrite(_nHandle,'	       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>	'+ c_ent)
fWrite(_nHandle,'	      </Borders>	'+ c_ent)
fWrite(_nHandle,'	      <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>	'+ c_ent)
fWrite(_nHandle,'	      <Interior/>	'+ c_ent)
fWrite(_nHandle,'	      <NumberFormat ss:Format="@"/>	'+ c_ent)
fWrite(_nHandle,'	     </Style>	'+ c_ent)
fWrite(_nHandle,'	    </Styles>	'+ c_ent)
fWrite(_nHandle,'	 <Worksheet ss:Name="1-Base extracao">	'+ c_ent)
fWrite(_nHandle,'	  <Names>	'+ c_ent)
fWrite(_nHandle,'	   <NamedRange ss:Name="_FilterDatabase"	'+ c_ent)
fWrite(_nHandle,'	    ss:RefersTo="='+"'1-Base extracao'"+'!R7C2:R8C29" ss:Hidden="1"/>	'+ c_ent)
fWrite(_nHandle,'	   <NamedRange ss:Name="Print_Area" ss:RefersTo="='+"'1-Base extracao'"+'!R7C3:R8C33"/>	'+ c_ent)
fWrite(_nHandle,'	  </Names>	'+ c_ent)
fWrite(_nHandle,'  <Table ss:ExpandedColumnCount="30" ss:ExpandedRowCount="10000" x:FullColumns="1"	'+ c_ent)  //Aqui
fWrite(_nHandle,'   x:FullRows="1" ss:StyleID="s63" ss:DefaultRowHeight="15">	'+ c_ent)
fWrite(_nHandle,'   <Column ss:StyleID="s63" ss:AutoFitWidth="0" ss:Width="10.5"/>	'+ c_ent)
fWrite(_nHandle,'   <Column ss:StyleID="s63" ss:AutoFitWidth="0" ss:Width="87.75"/>	'+ c_ent)
fWrite(_nHandle,'   <Column ss:StyleID="s63" ss:AutoFitWidth="0" ss:Width="106.5"/>	'+ c_ent)
fWrite(_nHandle,'   <Column ss:StyleID="s63" ss:AutoFitWidth="0" ss:Width="54"/>	'+ c_ent)
fWrite(_nHandle,'   <Column ss:StyleID="s63" ss:AutoFitWidth="0" ss:Width="66.75"/>	'+ c_ent)
fWrite(_nHandle,'   <Column ss:StyleID="s63" ss:Width="140.25"/>	'+ c_ent)
fWrite(_nHandle,'   <Column ss:StyleID="s63" ss:Width="141.75"/>	'+ c_ent)
fWrite(_nHandle,'   <Column ss:StyleID="s63" ss:Width="96"/>	'+ c_ent)
fWrite(_nHandle,'   <Column ss:StyleID="s63" ss:Width="206.25"/>	'+ c_ent)
fWrite(_nHandle,'   <Column ss:StyleID="s63" ss:AutoFitWidth="0" ss:Width="87.75"/>	'+ c_ent)
fWrite(_nHandle,'   <Column ss:StyleID="s63" ss:Width="125.25"/>	'+ c_ent)
fWrite(_nHandle,'   <Column ss:StyleID="s63" ss:Width="81.75"/>	'+ c_ent)
fWrite(_nHandle,'   <Column ss:StyleID="s63" ss:Width="91.5"/>	'+ c_ent)
fWrite(_nHandle,'   <Column ss:StyleID="s63" ss:Width="233.25"/>	'+ c_ent)
fWrite(_nHandle,'   <Column ss:StyleID="s64" ss:Width="107.25"/>	'+ c_ent)
fWrite(_nHandle,'   <Column ss:StyleID="s64" ss:Width="221.25"/>	'+ c_ent)
fWrite(_nHandle,'   <Column ss:StyleID="s64" ss:Width="132"/>	'+ c_ent)
fWrite(_nHandle,'   <Column ss:StyleID="s64" ss:Width="221.25"/>	'+ c_ent)
fWrite(_nHandle,'   <Column ss:StyleID="s64" ss:AutoFitWidth="0" ss:Width="107.25"/>	'+ c_ent)
fWrite(_nHandle,'   <Column ss:StyleID="s64" ss:Width="144.75"/>	'+ c_ent)
fWrite(_nHandle,'   <Column ss:StyleID="s64" ss:Width="144.75"/>	'+ c_ent) // Aqui
fWrite(_nHandle,'   <Column ss:StyleID="s63" ss:AutoFitWidth="0" ss:Width="92.25" ss:Span="7"/>	'+ c_ent)
fWrite(_nHandle,'   <Column ss:Index="30" ss:StyleID="s63" ss:AutoFitWidth="0" ss:Width="221.25"/>	'+ c_ent) //Aqui
fWrite(_nHandle,'   <Row ss:AutoFitHeight="0"/>	'+ c_ent)
fWrite(_nHandle,'   <Row ss:AutoFitHeight="0">	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:Index="2" ss:StyleID="s65"><Data ss:Type="String">Oi Futuro</Data></Cell>	'+ c_ent)
fWrite(_nHandle,'	   </Row>	'+ c_ent)
fWrite(_nHandle,'	   <Row ss:AutoFitHeight="0">	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:Index="2" ss:StyleID="s65"><Data ss:Type="String">Diretoria de Planejamento e Desempenho</Data></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s66"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s66"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s67"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s66"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s66"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s66"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s66"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s66"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:Index="15" ss:StyleID="s68"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s68"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s68"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s68"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s68"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s68"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s69"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s69"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s69"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s69"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s69"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s69"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s69"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s69"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s70"/>	'+ c_ent)
fWrite(_nHandle,'	   </Row>	'+ c_ent)
fWrite(_nHandle,'	   <Row ss:AutoFitHeight="0">	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:Index="2" ss:StyleID="s65"><Data ss:Type="String">Estrutura Oi Futuro</Data></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s66"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s66"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s67"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s66"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s66"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s66"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s66"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s66"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:Index="15" ss:StyleID="s68"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s68"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s68"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s68"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s68"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s68"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s69"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s69"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s69"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s69"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s69"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s69"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s69"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s69"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s70"/>	'+ c_ent)
fWrite(_nHandle,'	   </Row>	'+ c_ent)
fWrite(_nHandle,'	   <Row ss:AutoFitHeight="0">	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:Index="2" ss:StyleID="s66"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s66"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s66"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s67"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s66"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s66"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s66"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s66"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s66"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:Index="15" ss:StyleID="s68"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s68"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s68"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s68"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s68"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s68"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s69"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s69"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s69"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s69"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s69"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s69"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s69"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s69"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s70"/>	'+ c_ent)
fWrite(_nHandle,'	   </Row>	'+ c_ent)
fWrite(_nHandle,'	   <Row ss:AutoFitHeight="0" ss:StyleID="s71">	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:Index="7" ss:StyleID="s66"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:Index="10" ss:StyleID="s73"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s73"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s74"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:Index="15" ss:StyleID="s68"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s68"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s68"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s68"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s68"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s68"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s75"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s75"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s75"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s75"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s75"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s75"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s75"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s75"/>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s69"/>	'+ c_ent)
fWrite(_nHandle,'	   </Row>	'+ c_ent)           
fWrite(_nHandle,'	   <Row ss:AutoFitHeight="0">	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:Index="2" ss:StyleID="s76"><Data ss:Type="String">Tipos de Saldo</Data><NamedCell	'+ c_ent)
fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s76"><Data ss:Type="String">Classe orcamentaria</Data><NamedCell	'+ c_ent)
fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s76"><Data ss:Type="String">Mes</Data><NamedCell	'+ c_ent)
fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s76"><Data ss:Type="String">Data</Data><NamedCell	'+ c_ent)
fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s76"><Data ss:Type="String">Item contabil analitico</Data><NamedCell	'+ c_ent)
fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s76"><Data ss:Type="String">Item contabil sintetico</Data><NamedCell	'+ c_ent)
fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s76"><Data ss:Type="String">Centro de Custo</Data><NamedCell	'+ c_ent)
fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s76"><Data ss:Type="String">Operacao Orcamentaria</Data><NamedCell	'+ c_ent)
fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s76"><Data ss:Type="String">Nº doc fiscal</Data><NamedCell	'+ c_ent)
fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s76"><Data ss:Type="String">Data Emissao doc fiscal</Data><NamedCell	'+ c_ent)
fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s76"><Data ss:Type="String">Numero</Data><NamedCell	'+ c_ent)
fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s76"><Data ss:Type="String">Cod Fornecedor</Data><NamedCell	'+ c_ent)
fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s76"><Data ss:Type="String">Descricao Fornecedor</Data><NamedCell	'+ c_ent)
fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s76"><Data ss:Type="String">Cod Conta Contabil</Data><NamedCell	'+ c_ent)
fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s76"><Data ss:Type="String">Descricao conta contabil</Data><NamedCell	'+ c_ent)
fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s76"><Data ss:Type="String">Cod Conta Orcamentaria</Data><NamedCell	'+ c_ent)
fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s76"><Data ss:Type="String">Descricao conta Orcamentaria</Data><NamedCell	'+ c_ent)
fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s76"><Data ss:Type="String">Produto</Data><NamedCell	'+ c_ent)
fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s76"><Data ss:Type="String">Natureza</Data><NamedCell	'+ c_ent)
fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
//////////////////////////////////////////////////////////////////////////////////////////////////////////////
fWrite(_nHandle,'	    <Cell ss:StyleID="s76"><Data ss:Type="String">Valor PCO</Data><NamedCell	'+ c_ent)
fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)                                                                                            
//////////////////////////////////////////////////////////////////////////////////////////////////////////////
fWrite(_nHandle,'	    <Cell ss:StyleID="s76"><Data ss:Type="String">Valor</Data><NamedCell	'+ c_ent)
fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s76"><Data ss:Type="String">Valor Pago</Data><NamedCell	'+ c_ent)
fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s76"><Data ss:Type="String">PIS</Data><NamedCell	'+ c_ent)
fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s76"><Data ss:Type="String">COFINS</Data><NamedCell	'+ c_ent)
fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s76"><Data ss:Type="String">CSLL</Data><NamedCell	'+ c_ent)
fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s76"><Data ss:Type="String">ISS</Data><NamedCell	'+ c_ent)
fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s76"><Data ss:Type="String">IR</Data><NamedCell	'+ c_ent)
fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s76"><Data ss:Type="String">INSS</Data><NamedCell	'+ c_ent)
fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s76"><Data ss:Type="String">Descricao</Data><NamedCell	'+ c_ent)
fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	   </Row>	'+ c_ent)   

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////  
_cTotal := Len(_cArray3)  
ProcRegua(_cTotal)

For _ii := 1 to Len(_cArray3)

_cConta := _cConta + 1 
		
IncProc("Gerando Planilha - Linha: "+Alltrim(str(_cConta)) +" de "+ Alltrim(str(_cTotal)) +" ...")    

fWrite(_nHandle,'	   <Row ss:AutoFitHeight="0">	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:Index="2" ss:StyleID="s77"><Data ss:Type="String">'+Alltrim(_cArray3[_ii][1])+'</Data></Cell>'+ c_ent)
//fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s77"><Data ss:Type="String">'+Alltrim(_cArray3[_ii][2])+'</Data></Cell>'+ c_ent)
//fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s77"><Data ss:Type="String">'+Alltrim(_cArray3[_ii][3])+'</Data></Cell>'+ c_ent)
//fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s78"><Data ss:Type="String">'+DTOC(STOD(_cArray3[_ii][4]))+'</Data></Cell>'+ c_ent)
//fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s77"><Data ss:Type="String">'+Alltrim(_cArray3[_ii][5])+'</Data></Cell>'+ c_ent)
//fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s77"><Data ss:Type="String">'+Alltrim(_cArray3[_ii][6])+'</Data></Cell>'+ c_ent)
//fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s77"><Data ss:Type="String">'+Alltrim(_cArray3[_ii][7])+'</Data></Cell>'+ c_ent)
//fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s77"><Data ss:Type="String">'+Alltrim(_cArray3[_ii][8])+'</Data></Cell>'+ c_ent)
//fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s77"><Data ss:Type="String">'+Alltrim(_cArray3[_ii][9])+'</Data></Cell>'+ c_ent)
//fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s77"><Data ss:Type="String">'+Alltrim(_cArray3[_ii][10])+'</Data></Cell>'+ c_ent)
//fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s79"><Data ss:Type="String">'+Alltrim(_cArray3[_ii][11])+'</Data></Cell>'+ c_ent)
//fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s77"><Data ss:Type="String">'+Alltrim(_cArray3[_ii][12])+'</Data></Cell>'+ c_ent)
//fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)              
fWrite(_nHandle,'	    <Cell ss:StyleID="s77"><Data ss:Type="String">'+Alltrim(_cArray3[_ii][13])+'</Data></Cell>'+ c_ent)
//fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s79"><Data ss:Type="String">'+Alltrim(_cArray3[_ii][14])+'</Data></Cell>'+ c_ent)
//fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s79"><Data ss:Type="String">'+Alltrim(_cArray3[_ii][15])+'</Data></Cell>'+ c_ent)
//fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s79"><Data ss:Type="String">'+Alltrim(_cArray3[_ii][16])+'</Data></Cell>'+ c_ent)
//fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s79"><Data ss:Type="String">'+Alltrim(_cArray3[_ii][17])+'</Data></Cell>'+ c_ent)
//fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s79"><Data ss:Type="String">'+Alltrim(_cArray3[_ii][18])+'</Data></Cell>'+ c_ent)
//fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s79"><Data ss:Type="String">'+Alltrim(_cArray3[_ii][19])+'</Data></Cell>'+ c_ent)
//fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)      
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
fWrite(_nHandle,'	    <Cell ss:StyleID="s81"><Data ss:Type="String">'+Strtran(Alltrim(Transform(IIf(_cArray3[_ii][29] == "1",_cArray3[_ii][30],Iif(!Empty(_cArray3[_ii][30]),_cArray3[_ii][30]*-1,_cArray3[_ii][30])), "@E 999999999.99" )),",",",")+'</Data></Cell>'+ c_ent)
//fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
fWrite(_nHandle,'	    <Cell ss:StyleID="s81"><Data ss:Type="String">'+Strtran(Alltrim(Transform(_cArray3[_ii][20], "@E 999999999.99" )),",",",")+'</Data></Cell>'+ c_ent)
//fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s81"><Data ss:Type="String">'+Strtran(Alltrim(Transform(_cArray3[_ii][21], "@E 999999999.99" )),",",",")+'</Data></Cell>'+ c_ent)
//fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s81"><Data ss:Type="String">'+Strtran(Alltrim(Transform(_cArray3[_ii][22], "@E 999999999.99" )),",",",")+'</Data></Cell>'+ c_ent)
//fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s81"><Data ss:Type="String">'+Strtran(Alltrim(Transform(_cArray3[_ii][23], "@E 999999999.99" )),",",",")+'</Data></Cell>'+ c_ent)
//fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s81"><Data ss:Type="String">'+Strtran(Alltrim(Transform(_cArray3[_ii][24], "@E 999999999.99" )),",",",")+'</Data></Cell>'+ c_ent)
//fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s81"><Data ss:Type="String">'+Strtran(Alltrim(Transform(_cArray3[_ii][25], "@E 999999999.99" )),",",",")+'</Data></Cell>'+ c_ent)
//fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s81"><Data ss:Type="String">'+Strtran(Alltrim(Transform(_cArray3[_ii][26], "@E 999999999.99" )),",",",")+'</Data></Cell>'+ c_ent)
//fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s81"><Data ss:Type="String">'+Strtran(Alltrim(Transform(_cArray3[_ii][27], "@E 999999999.99" )),",",",")+'</Data></Cell>'+ c_ent)
//fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	    <Cell ss:StyleID="s79"><Data ss:Type="String">'+Alltrim(_cArray3[_ii][28])+'</Data></Cell>'+ c_ent)  
//fWrite(_nHandle,'	      ss:Name="_FilterDatabase"/></Cell>	'+ c_ent)
fWrite(_nHandle,'	   </Row>	'+ c_ent)                                          

   /*	fWrite(_nHandle,'   <Row ss:AutoFitHeight="0">'+ c_ent)
	fWrite(_nHandle,'    <Cell ss:Index="2" ss:StyleID="s162"><Data ss:Type="String">'+_cArray3[_ii][1]+'</Data></Cell>'+ c_ent)
	fWrite(_nHandle,'    <Cell ss:StyleID="s162"><Data ss:Type="String">'+_cArray3[_ii][2]+'</Data></Cell>'+ c_ent)
	fWrite(_nHandle,'    <Cell ss:StyleID="s162"><Data ss:Type="String">'+_cArray3[_ii][3]+'</Data></Cell>'+ c_ent)
	fWrite(_nHandle,'    <Cell ss:StyleID="s166"><Data ss:Type="String">'+DTOC(STOD(_cArray3[_ii][4]))+'</Data></Cell>'+ c_ent)
	fWrite(_nHandle,'    <Cell ss:StyleID="s162"><Data ss:Type="String">'+_cArray3[_ii][5]+'</Data></Cell>'+ c_ent)
	fWrite(_nHandle,'    <Cell ss:StyleID="s162"><Data ss:Type="String">'+_cArray3[_ii][6]+'</Data></Cell>'+ c_ent)
	fWrite(_nHandle,'    <Cell ss:StyleID="s162"><Data ss:Type="String">'+_cArray3[_ii][7]+'</Data></Cell>'+ c_ent)
	fWrite(_nHandle,'    <Cell ss:StyleID="s162"><Data ss:Type="String">'+_cArray3[_ii][8]+'</Data></Cell>'+ c_ent)
	fWrite(_nHandle,'    <Cell ss:StyleID="s162"><Data ss:Type="String">'+_cArray3[_ii][9]+'</Data></Cell>'+ c_ent)
	fWrite(_nHandle,'    <Cell ss:StyleID="s162"><Data ss:Type="String">'+_cArray3[_ii][10]+'</Data></Cell>'+ c_ent)
	fWrite(_nHandle,'    <Cell ss:StyleID="s164"><Data ss:Type="String">'+_cArray3[_ii][11]+'</Data></Cell>'+ c_ent)
	fWrite(_nHandle,'    <Cell ss:StyleID="s162"><Data ss:Type="String">'+_cArray3[_ii][12]+'</Data></Cell>'+ c_ent)
	fWrite(_nHandle,'    <Cell ss:StyleID="s162"><Data ss:Type="String">'+_cArray3[_ii][13]+'</Data></Cell>'+ c_ent)
	fWrite(_nHandle,'    <Cell ss:StyleID="s164"><Data ss:Type="String">'+_cArray3[_ii][14]+'</Data></Cell>'+ c_ent)
	fWrite(_nHandle,'    <Cell ss:StyleID="s164"><Data ss:Type="String">'+_cArray3[_ii][15]+'</Data></Cell>'+ c_ent)
	fWrite(_nHandle,'    <Cell ss:StyleID="s164"><Data ss:Type="String">'+_cArray3[_ii][16]+'</Data></Cell>'+ c_ent)
	fWrite(_nHandle,'    <Cell ss:StyleID="s164"><Data ss:Type="String">'+_cArray3[_ii][17]+'</Data></Cell>'+ c_ent)
	fWrite(_nHandle,'    <Cell ss:StyleID="s164"><Data ss:Type="String">'+_cArray3[_ii][18]+'</Data></Cell>'+ c_ent)
	fWrite(_nHandle,'    <Cell ss:StyleID="s164"><Data ss:Type="String">'+_cArray3[_ii][19]+'</Data></Cell>'+ c_ent)
	fWrite(_nHandle,'    <Cell ss:StyleID="s165"><Data ss:Type="String">'+_cArray3[_ii][20]+'</Data></Cell>'+ c_ent)
	fWrite(_nHandle,'    <Cell ss:StyleID="s165"><Data ss:Type="String">'+Strtran(Alltrim(Transform(_cArray3[_ii][21], "@E 999999999.99" )),",",".")+'</Data></Cell>'+ c_ent)
	fWrite(_nHandle,'    <Cell ss:StyleID="s165"><Data ss:Type="String">'+Strtran(Alltrim(Transform(_cArray3[_ii][22], "@E 999999999.99" )),",",".")+'</Data></Cell>'+ c_ent)
	fWrite(_nHandle,'    <Cell ss:StyleID="s165"><Data ss:Type="String">'+Strtran(Alltrim(Transform(_cArray3[_ii][23], "@E 999999999.99" )),",",".")+'</Data></Cell>'+ c_ent)
	fWrite(_nHandle,'    <Cell ss:StyleID="s165"><Data ss:Type="String">'+Strtran(Alltrim(Transform(_cArray3[_ii][24], "@E 999999999.99" )),",",".")+'</Data></Cell>'+ c_ent)
	fWrite(_nHandle,'    <Cell ss:StyleID="s165"><Data ss:Type="String">'+Strtran(Alltrim(Transform(_cArray3[_ii][25], "@E 999999999.99" )),",",".")+'</Data></Cell>'+ c_ent)
	fWrite(_nHandle,'    <Cell ss:StyleID="s165"><Data ss:Type="String">'+Strtran(Alltrim(Transform(_cArray3[_ii][26], "@E 999999999.99" )),",",".")+'</Data></Cell>'+ c_ent)
	fWrite(_nHandle,'    <Cell ss:StyleID="s165"><Data ss:Type="String">'+_cArray3[_ii][27]+'</Data></Cell>'+ c_ent)
	fWrite(_nHandle,'    <Cell ss:StyleID="s164"><Data ss:Type="String">'+_cArray3[_ii][28]+'</Data></Cell>'+ c_ent)
	fWrite(_nHandle,'   </Row>'+ c_ent)  */ 
	
Next _ii

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

fWrite(_nHandle,'	  </Table>	'+ c_ent)
fWrite(_nHandle,'	  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">	'+ c_ent)
fWrite(_nHandle,'	   <PageSetup>	'+ c_ent)
fWrite(_nHandle,'	    <Layout x:Orientation="Landscape"/>	'+ c_ent)
fWrite(_nHandle,'	    <Header x:Margin="0.3"/>	'+ c_ent)
fWrite(_nHandle,'	    <Footer x:Margin="0.3"/>	'+ c_ent)
fWrite(_nHandle,'	    <PageMargins x:Bottom="0.75" x:Left="0.25" x:Right="0.25" x:Top="0.75"/>	'+ c_ent)
fWrite(_nHandle,'	   </PageSetup>	'+ c_ent)
fWrite(_nHandle,'	   <Unsynced/>	'+ c_ent)
fWrite(_nHandle,'	   <FitToPage/>	'+ c_ent)
fWrite(_nHandle,'	   <Print>	'+ c_ent)
fWrite(_nHandle,'	    <ValidPrinterInfo/>	'+ c_ent)
fWrite(_nHandle,'	    <PaperSizeIndex>9</PaperSizeIndex>	'+ c_ent)
fWrite(_nHandle,'	    <Scale>34</Scale>	'+ c_ent)
fWrite(_nHandle,'	    <HorizontalResolution>600</HorizontalResolution>	'+ c_ent)
fWrite(_nHandle,'	    <VerticalResolution>600</VerticalResolution>	'+ c_ent)
fWrite(_nHandle,'	   </Print>	'+ c_ent)
fWrite(_nHandle,'	   <TabColorIndex>53</TabColorIndex>	'+ c_ent)
fWrite(_nHandle,'	   <Selected/>	'+ c_ent)
fWrite(_nHandle,'	   <DoNotDisplayGridlines/>	'+ c_ent)
//fWrite(_nHandle,'	   <LeftColumnVisible>1</LeftColumnVisible>	'+ c_ent)
fWrite(_nHandle,'	   <Panes>	'+ c_ent)
fWrite(_nHandle,'	    <Pane>	'+ c_ent)
fWrite(_nHandle,'	     <Number>3</Number>	'+ c_ent)
fWrite(_nHandle,'	     <ActiveRow>1</ActiveRow>	'+ c_ent)
//fWrite(_nHandle,'	     <ActiveCol>20</ActiveCol>	'+ c_ent)
fWrite(_nHandle,'	    </Pane>	'+ c_ent)
fWrite(_nHandle,'	   </Panes>	'+ c_ent)
fWrite(_nHandle,'	   <ProtectObjects>False</ProtectObjects>	'+ c_ent)
fWrite(_nHandle,'	   <ProtectScenarios>False</ProtectScenarios>	'+ c_ent)
fWrite(_nHandle,'	  </WorksheetOptions>	'+ c_ent)
fWrite(_nHandle,'	  <AutoFilter x:Range="R7C2:R8C29"	'+ c_ent)
fWrite(_nHandle,'	   xmlns="urn:schemas-microsoft-com:office:excel">	'+ c_ent)
fWrite(_nHandle,'	  </AutoFilter>	'+ c_ent)
fWrite(_nHandle,'	 </Worksheet>	'+ c_ent)
fWrite(_nHandle,'	</Workbook>	'+ c_ent)


//****************|
//Fechando arquivo|
//****************|
fClose(_nHandle)

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³Abre a Planilha Excel³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
ShellExecute( "Open", _cArquivo, " ", _cPath, 3 )

RestArea(_cOIFR004)

Return()  

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³OIFR004   ºAutor  ³Leandro Ribeiro     º Data ³  06/21/14   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Função para retornar a descrição das contas sinteticas.    º±±
±±º          ³ Criada a função pois as funções GetAdvfval e Posicione não º±±
±±º          ³ funcionava.												  º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP                                                        º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function RETITSINT(_cContSit)

Local _cArea4  := GetArea()
Local _cQuery4 := ""
Local _cTrab4  := GetNextAlias() 
Local _cArray4 := {}

_cQuery4 := " SELECT CTD_DESC01 "+ c_ent
_cQuery4 += " FROM "+RETSQLNAME("CTD")+" CTD "+ c_ent
_cQuery4 += " WHERE "+ c_ent
_cQuery4 += " CTD_FILIAL = '"+xFilial("CTD")+"' "+ c_ent
_cQuery4 += " AND CTD_ITEM = (SELECT CTD_ITSUP "+ c_ent
_cQuery4 += " FROM "+RETSQLNAME("CTD")+" CTD "+ c_ent
_cQuery4 += " WHERE "+ c_ent 
_cQuery4 += " CTD_FILIAL = '"+xFilial("CTD")+"'"+ c_ent
_cQuery4 += " AND CTD_ITEM = '"+_cContSit+"'"+ c_ent
_cQuery4 += " AND CTD.D_E_L_E_T_ = ' ')"+ c_ent
_cQuery4 += " AND CTD.D_E_L_E_T_ = ' '"+ c_ent 
_cQuery4 := ChangeQuery(_cQuery4)   

If Select(_cTrab4) > 0
	dbSelectArea(_cTrab4)
	(_cTrab4)->(dbCloseArea())
EndIf

dbUseArea(.T.,"TOPCONN",TcGenQry(,,_cQuery4),_cTrab4,.T.,.T.)       

dbSelectArea(_cTrab4)

If(!Eof())
  While !(_cTrab4)->(Eof())  
		Aadd(_cArray4,AllTrim((_cTrab4)->CTD_DESC01))
  	DbSkip()
  EndDo 
EndIf  

(_cTrab4)->(dbCloseArea())    

RestArea(_cArea4)

Return(_cArray4)   

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³OIFR004   ºAutor  ³Leandro Ribeiro     º Data ³  06/21/14   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Função para filtrar os fornecedores de acordo com os para- º±±
±±º          ³ metros.                                                    º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP                                                        º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function RETFORNECE(_cForn1,_cForn2)

Local _cArea1  := GetArea()
Local _cQuery1 := ""
Local _cTrab1  := GetNextAlias() 
Local _cArray1 := {}

_cQuery1 := " SELECT DISTINCT A2_COD "+ c_ent
_cQuery1 += " FROM "+RETSQLNAME("SA2")+" SA2 "+ c_ent
_cQuery1 += " WHERE "+ c_ent
_cQuery1 += " A2_FILIAL = '"+xFilial("SA2")+"'"+ c_ent
_cQuery1 += " AND A2_COD BETWEEN '"+Alltrim(_cForn1)+"' AND '"+Alltrim(_cForn2)+"' "+ c_ent
_cQuery1 += " AND SA2.D_E_L_E_T_ = ' ' "+ c_ent
_cQuery1 := ChangeQuery(_cQuery1)   

If Select(_cTrab1) > 0
	dbSelectArea(_cTrab1)
	(_cTrab1)->(dbCloseArea())
EndIf

dbUseArea(.T.,"TOPCONN",TcGenQry(,,_cQuery1),_cTrab1,.T.,.T.)       

dbSelectArea(_cTrab1)

If(!Eof())
  While !(_cTrab1)->(Eof())  
		Aadd(_cArray1,(_cTrab1)->A2_COD)
  	DbSkip()
  EndDo 
EndIf  

(_cTrab1)->(dbCloseArea())    

RestArea(_cArea1)

Return(_cArray1)  

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³OIFR004   ºAutor  ³Leandro Ribeiro     º Data ³  06/21/14   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Função para retornar as contas contabeis                   º±±
±±º          ³ Função de continuada.                                      º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP                                                        º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function RETCONTCTB(_cCtCTB1,_cCtCTB2)

Local _cArea2  := GetArea()
Local _cQuery2 := ""
Local _cTrab2  := GetNextAlias() 
Local _cArray2 := {}

_cQuery2 := " SELECT DISTINCT CT1_CONTA "+ c_ent
_cQuery2 += " FROM "+RETSQLNAME("CT1")+" CT1 "+ c_ent
_cQuery2 += " WHERE "+ c_ent    
_cQuery2 += " CT1_FILIAL = '"+xFilial("CT1")+"'"+ c_ent    
_cQuery2 += " AND CT1_CONTA BETWEEN '"+Alltrim(_cCtCTB1)+"' AND '"+Alltrim(_cCtCTB2)+"' "+ c_ent
_cQuery2 += " AND CT1.D_E_L_E_T_ = ' ' "+ c_ent
_cQuery2 := ChangeQuery(_cQuery2)   

If Select(_cTrab2) > 0
	dbSelectArea(_cTrab2)
	(_cTrab2)->(dbCloseArea())
EndIf

dbUseArea(.T.,"TOPCONN",TcGenQry(,,_cQuery2),_cTrab2,.T.,.T.)       

dbSelectArea(_cTrab2)

If(!Eof())
  While !(_cTrab2)->(Eof())  
		Aadd(_cArray2,(_cTrab2)->CT1_CONTA)
  	DbSkip()
  EndDo 
EndIf  

(_cTrab2)->(dbCloseArea())    

RestArea(_cArea2)

Return(_cArray2)         


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³OIFR004   ºAutor  ³Leandro Ribeiro     º Data ³  06/21/14   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Função para procurar os registros de acordo com a chave    º±±
±±º          ³ gravada na AKD.                                            º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP                                                        º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function RETCHAVE(_cChave,_xForn)

Local _cArea5  := GetArea()
Local _cAlias  := SUBSTR(_cChave,1,3) 
Local _cIndice := SUBSTR(_cChave,4)   
Local _cRDados := {}     
Local _lFlag   := .F.   
Local _cQuery3 := ""
Local _cTrab3  := GetNextAlias() 
Local _cNaturez := ""
Local _cVlPago  := ""
Local _cNumDoc := ""

If(_cAlias == "ZZA") 
	DbSelectArea(_cAlias)
	DbSetOrder(1)
	If(DbSeek(_cIndice))
		_cAlias  := ZZA->ZZA_TPORIG	
		_cIndice := ZZA->ZZA_E2X2UN 
		If(_cAlias $ "SD1|SDE")
			_cNumDoc := ZZA->ZZA_NUMERO
		EndIF
		_lFlag := .T.
	EndIf
EndIf

Do Case
	Case(_cAlias == "SC7")              
				
		DbSelectArea(_cAlias)
		DbSetOrder(1)
		If(DbSeek(_cIndice))
					 
			For _rr := 1 to Len(_xForn)
				
				If(SC7->C7_FORNECE == _xForn[_rr])		
				
					Aadd(_cRDados,{"",; //NUMERO DA NOTA FISCAL
								   "",; //SERIE DA NOTA FISCAL
					SC7->C7_NUM,; //NUMERO DO PEDIDO DE COMPRAS
					SC7->C7_FORNECE,; //CODIGO DO FORNECEDOR
					TiraGraf(Alltrim(POSICIONE("SA2",1,xFilial("SA2")+PADR(SC7->C7_FORNECE,TAMSX3("C7_FORNECE")[1]),"A2_NOME"))),; //NOME DO FORNECEDOR
					SC7->C7_CONTA,; //CONTA CONTABIL
					TiraGraf(Alltrim(POSICIONE("CT1",1,xFilial("CT1")+PADR(SC7->C7_CONTA,TAMSX3("C7_CONTA")[1]),"CT1_DESC01"))),; //DESCRIÇÃO CONTA CONTABIL
					TiraGraf(Alltrim(POSICIONE("SB1",1,xFilial("SB1")+PADR(SC7->C7_PRODUTO,TAMSX3("C7_PRODUTO")[1]),"B1_DESC"))),; //PRODUTO
					"",; //NATUREZA
					SC7->C7_PRECO,; //VALOR
					"",; //VALOR PAGO
					"",; //PIS
					"",; //COFINS
					"",; //CSLL
					"",; //ISS
					SC7->C7_VALIR,; //IR
					"",; //INSS
					Alltrim(SC7->C7_DESCRI)}) //MENSAGEM    
				
				EndIf
				
	   		Next _rr
			(_cAlias)->(DbCloseArea())	
		EndIf
		
	Case(_cAlias == "SE2")
	
		DbSelectArea(_cAlias)
		DbSetOrder(1)
		If(DbSeek(_cIndice))      
		
			For _rr := 1 to Len(_xForn)
				
				If(SE2->E2_FORNECE == _xForn[_rr])
		
					//Aadd(_cRDados,{"",; //NUMERO DA NOTA FISCAL
					Aadd(_cRDados,{SE2->E2_NUM,; //NUMERO DA NOTA FISCAL
					SE2->E2_PREFIXO,; //SERIE DA NOTA FISCAL
					SE2->E2_NUM,; //NUMERO DO PEDIDO DE COMPRAS
					SE2->E2_FORNECE,; //CODIGO DO FORNECEDOR
					TiraGraf(Alltrim(POSICIONE("SA2",1,xFilial("SA2")+PADR(SE2->E2_FORNECE,TAMSX3("E2_FORNECE")[1]),"A2_NOME"))),; //NOME DO FORNECEDOR
					SE2->E2_CONTAD,; //CONTA CONTABIL
					TiraGraf(Alltrim(POSICIONE("CT1",1,xFilial("CT1")+PADR(SE2->E2_CONTAD,TAMSX3("D1_CONTA")[1]),"CT1_DESC01"))),; //DESCRIÇÃO CONTA CONTABIL
					"",; //PRODUTO
					SE2->E2_NATUREZ,; //NATUREZA
					SE2->E2_VALOR,; //VALOR
					SE2->E2_VALLIQ,; //VALOR PAGO
					SE2->E2_PIS,; //PIS
					SE2->E2_COFINS,; //COFINS
					SE2->E2_CSLL,; //CSLL
					SE2->E2_ISS,; //ISS
					SE2->E2_IRRF,; //IR
					SE2->E2_INSS,; //INSS
					Alltrim(SE2->E2_HIST)}) //MENSAGEM 
				
				EndIf
				
			Next _rr
			(_cAlias)->(DbCloseArea())
		EndIf		
		
	Case(_cAlias == "SE5") 	                     
		
			DbSelectArea(_cAlias)
			DbSetOrder(2)
			If(DbSeek(_cIndice))   		  
		
			For _rr := 1 to Len(_xForn)
				
				If(SE5->E5_CLIFOR == _xForn[_rr])
		
					Aadd(_cRDados,{SE5->E5_NUMERO,; //NUMERO DA NOTA FISCAL
					SE5->E5_PREFIXO,; //SERIE DA NOTA FISCAL
					SE5->E5_NUMERO,; //NUMERO DO PEDIDO DE COMPRAS
					SE5->E5_CLIFOR,; //CODIGO DO FORNECEDOR
					SE5->E5_BENEF,; //NOME DO FORNECEDOR
					"",; //CONTA CONTABIL
					"",; //DESCRIÇÃO CONTA CONTABIL
					"",; //PRODUTO
					SE5->E5_NATUREZ,; //NATUREZA
					SE5->E5_VALOR,; //VALOR
					SE5->E5_VALOR,; //VALOR PAGO
					SE5->E5_VRETPIS,; //PIS
					SE5->E5_VRETCOF,; //COFINS
					SE5->E5_VRETCSL,; //CSLL
					SE5->E5_VRETISS,; //ISS
					SE5->E5_VRETIRF,; //IR
					"",; //INSS
					SE5->E5_HISTOR}) //MENSAGEM	 
					
				EndIf
				
			Next _rr
	   		(_cAlias)->(DbCloseArea())
		EndIf	
	
	Case(_cAlias == "SD1")  
	    
	    ///////////////////////////////
		// Função para validar a ZZA //
		///////////////////////////////
		If(_lFlag)
			
			DbSelectArea("SE2")
			DbSetOrder(1)
			If(DbSeek(_cIndice)) 
				
				_cQuery3 := " SELECT * FROM "+RETSQLNAME("SD1")+" SD1 "+ c_ent
				_cQuery3 += " INNER JOIN "+RETSQLNAME("SE2")+" SE2 "+ c_ent
				_cQuery3 += " ON D1_FILIAL = E2_FILIAL"+ c_ent
				_cQuery3 += " AND D1_FORNECE = E2_FORNECE"+ c_ent
				_cQuery3 += " AND D1_LOJA = E2_LOJA"+ c_ent
				//_cQuery3 += " AND D1_TOTAL = E2_VALOR"+ c_ent 
				_cQuery3 += " AND D1_DOC = E2_NUM"+ c_ent
				_cQuery3 += " WHERE"+ c_ent
				_cQuery3 += " D1_FORNECE = '"+SE2->E2_FORNECE+"'
				_cQuery3 += " AND D1_LOJA = '"+SE2->E2_LOJA+"'   
				_cQuery3 +=	" AND D1_DOC = '"+_cNumDoc+"'
				_cQuery3 +=	" AND SD1.D_E_L_E_T_ = ''
				_cQuery3 +=	" AND SE2.D_E_L_E_T_ = ''
				//_cQuery3 += " AND D1_TOTAL = '"+CVALTOCHAR(SE2->E2_VALOR)+"'   
				_cQuery3 := ChangeQuery(_cQuery3)   
					
					If Select(_cTrab3) > 0
						dbSelectArea(_cTrab3)
						(_cTrab3)->(dbCloseArea())
					EndIf
					
					dbUseArea(.T.,"TOPCONN",TcGenQry(,,_cQuery3),_cTrab3,.T.,.T.)       
					
					dbSelectArea(_cTrab3)
					
					If(!Eof())
					  While !(_cTrab3)->(Eof())  
					    	_cIndice := (_cTrab3)->D1_FILIAL+(_cTrab3)->D1_DOC+(_cTrab3)->D1_SERIE+(_cTrab3)->D1_FORNECE+(_cTrab3)->D1_LOJA+(_cTrab3)->D1_COD+(_cTrab3)->D1_ITEM		
					  	DbSkip()
					  EndDo 
					EndIf 
					
					_cNaturez := SE2->E2_NATUREZ 
					_cVlPago  := SE2->E2_VALLIQ
					
					(_cTrab3)->(dbCloseArea()) 
			
			EndIf  
			
		EndIf	
		////////////////////////////////////////////////////////////
		
		DbSelectArea(_cAlias)
		DbSetOrder(1)
		If(DbSeek(_cIndice))  
			
			For _rr := 1 to Len(_xForn)
				
				If(SD1->D1_FORNECE == _xForn[_rr])
				
					Aadd(_cRDados,{SD1->D1_DOC,; //NUMERO DA NOTA FISCAL
					DTOC(SD1->D1_EMISSAO),; //DATA DE EMISSAO DA NOTA
					SD1->D1_PEDIDO,; //NUMERO DO PEDIDO DE COMPRAS
					SD1->D1_FORNECE,; //CODIGO DO FORNECEDOR
					TiraGraf(Alltrim(POSICIONE("SA2",1,xFilial("SA2")+PADR(SD1->D1_FORNECE,TAMSX3("D1_FORNECE")[1]),"A2_NOME"))),; //NOME DO FORNECEDOR
					SD1->D1_CONTA,; //CONTA CONTABIL
					TiraGraf(Alltrim(POSICIONE("CT1",1,xFilial("CT1")+PADR(SD1->D1_CONTA,TAMSX3("D1_CONTA")[1]),"CT1_DESC01"))),; //DESCRIÇÃO CONTA CONTABIL,; //DESCRIÇÃO CONTA CONTABIL
					TiraGraf(Alltrim(POSICIONE("SB1",1,xFilial("SB1")+PADR(SD1->D1_COD,TAMSX3("D1_COD")[1]),"B1_DESC"))),; //PRODUTO,; //PRODUTO
				    AllTrim(_cNaturez),; //NATUREZA
					SD1->D1_TOTAL,; //VALOR
					_cVlPago,; //VALOR PAGO
					SD1->D1_VALPIS,; //PIS
					SD1->D1_VALCOF,; //COFINS
					SD1->D1_VALCSL,; //CSLL
					SD1->D1_VALISS,; //ISS
					SD1->D1_VALIRR,; //IR
					SD1->D1_VALINS,; //INSS
					Alltrim(SD1->D1_DESCRI)}) //MENSAGEM 
					
				EndIf
			
			Next _rr   
			(_cAlias)->(DbCloseArea())
		EndIf 
		
	Case(_cAlias == "SCH")	
	
		DbSelectArea(_cAlias)
		DbSetOrder(1)
		If(DbSeek(_cIndice))
			For _rr := 1 to Len(_xForn)
				If(SCH->CH_FORNECE == _xForn[_rr])
				
					DbSelectArea("SC7")
					DbSetOrder(1)
					DbSeek(xFilial("SC7")+SCH->CH_PEDIDO+SCH->CH_ITEMPD)		
				
					Aadd(_cRDados,{"",; //NUMERO DA NOTA FISCAL
					"",; //SERIE DA NOTA FISCAL
					SCH->CH_PEDIDO,; //NUMERO DO PEDIDO DE COMPRAS
					SCH->CH_FORNECE,; //CODIGO DO FORNECEDOR
					TiraGraf(Alltrim(POSICIONE("SA2",1,xFilial("SA2")+PADR(SCH->CH_FORNECE,TAMSX3("CH_FORNECE")[1]),"A2_NOME"))),; //NOME DO FORNECEDOR  //PONTERAR
					SCH->CH_CONTA,; //CONTA CONTABIL
					TiraGraf(Alltrim(POSICIONE("CT1",1,xFilial("CT1")+PADR(SCH->CH_CONTA,TAMSX3("D1_CONTA")[1]),"CT1_DESC01"))),; //DESCRIÇÃO CONTA CONTABIL // PONTERAR
					TiraGraf(Alltrim(POSICIONE("SB1",1,xFilial("SB1")+PADR(SC7->C7_PRODUTO,TAMSX3("C7_PRODUTO")[1]),"B1_DESC"))),; //PRODUTO // PEDIDO DE COMPRAS
					"",; //NATUREZA
					(SC7->C7_PRECO *  SCH->CH_PERC) / 100,; //VALOR
					"",; //VALOR PAGO
					"",; //PIS
					"",; //COFINS
					"",; //CSLL
					"",; //ISS
					SC7->C7_VALIR,; //IR
					"",; //INSS
					Alltrim(SC7->C7_DESCRI)}) //MENSAGEM  				
				
				EndIf
		    Next _rr 
		    SC7->(DbCloseArea()) 
		    (_cAlias)->(DbCloseArea())				
		EndIf   
		
	Case(_cAlias == "SDE")	
		
		///////////////////////////////
		// Função para validar a ZZA //
		///////////////////////////////
		If(_lFlag)
			
			DbSelectArea("SE2")
			DbSetOrder(1)
			If(DbSeek(_cIndice)) 
				
				_cQuery3 := " SELECT D1_FILIAL,D1_DOC,D1_SERIE,D1_FORNECE,D1_LOJA,D1_COD,D1_ITEM,E2_NATUREZ,E2_VALLIQ FROM "+RETSQLNAME("SD1")+" SD1 "+ c_ent
				_cQuery3 += " INNER JOIN "+RETSQLNAME("SE2")+" SE2 "+ c_ent
				_cQuery3 += " ON D1_FILIAL = E2_FILIAL"+ c_ent
				_cQuery3 += " AND D1_FORNECE = E2_FORNECE"+ c_ent
				_cQuery3 += " AND D1_LOJA = E2_LOJA"+ c_ent
				//_cQuery3 += " AND D1_TOTAL = E2_VALOR"+ c_ent
				_cQuery3 += " AND D1_DOC = E2_NUM"+ c_ent
				_cQuery3 += " WHERE"+ c_ent
				_cQuery3 += " D1_FORNECE = '"+SE2->E2_FORNECE+"'"+ c_ent
				_cQuery3 += " AND D1_LOJA = '"+SE2->E2_LOJA+"'"+ c_ent    
				_cQuery3 +=	" AND D1_DOC = '"+_cNumDoc+"'
				//_cQuery3 += " AND D1_TOTAL = '"+CVALTOCHAR(SE2->E2_VALOR)+"'"+ c_ent ????  
				_cQuery3 +=	" AND SD1.D_E_L_E_T_ = ''
				_cQuery3 +=	" AND SE2.D_E_L_E_T_ = ''
				_cQuery3 += " GROUP BY D1_FILIAL,D1_DOC,D1_SERIE,D1_FORNECE,D1_LOJA,D1_COD,D1_ITEM,E2_NATUREZ,E2_VALLIQ "+ c_ent  
				_cQuery3 := ChangeQuery(_cQuery3)   
					
					If Select(_cTrab3) > 0
						dbSelectArea(_cTrab3)
						(_cTrab3)->(dbCloseArea())
					EndIf
					
					dbUseArea(.T.,"TOPCONN",TcGenQry(,,_cQuery3),_cTrab3,.T.,.T.)       
					
					dbSelectArea(_cTrab3)
					
					If(!Eof())
					  While !(_cTrab3)->(Eof())  
					    	_cIndice := (_cTrab3)->D1_FILIAL+(_cTrab3)->D1_DOC+(_cTrab3)->D1_SERIE+(_cTrab3)->D1_FORNECE+(_cTrab3)->D1_LOJA+(_cTrab3)->D1_ITEM//(_cTrab3)->D1_COD+		
					    	_cNaturez := (_cTrab3)->E2_NATUREZ
					  	DbSkip()
					  EndDo 
					EndIf  
					
					(_cTrab3)->(dbCloseArea()) 
			
			EndIf  
			
		EndIf	
		////////////////////////////////////////////////////////////
	
		DbSelectArea(_cAlias)
		DbSetOrder(1)
		If(DbSeek(_cIndice))
			For _rr := 1 to Len(_xForn)
				If(SDE->DE_FORNECE == _xForn[_rr])
				
					DbSelectArea("SD1")
					DbSetOrder(1)
					DbSeek(xFilial("SD1")+SDE->DE_DOC+SDE->DE_SERIE+SDE->DE_FORNECE)//+SDE->DE_ITEMNF+SDE->DE_ITEM)		
				
					Aadd(_cRDados,{SD1->D1_DOC,; //NUMERO DA NOTA FISCAL
					DTOC(SD1->D1_EMISSAO),; //DATA DE EMISSAO DA NOTA
					SD1->D1_PEDIDO,; //NUMERO DO PEDIDO DE COMPRAS
					SD1->D1_FORNECE,; //CODIGO DO FORNECEDOR
					TiraGraf(Alltrim(POSICIONE("SA2",1,xFilial("SA2")+PADR(SD1->D1_FORNECE,TAMSX3("D1_FORNECE")[1]),"A2_NOME"))),; //NOME DO FORNECEDOR
					SDE->DE_CONTA,; //CONTA CONTABIL
					TiraGraf(Alltrim(POSICIONE("CT1",1,xFilial("CT1")+PADR(AllTrim(SDE->DE_CONTA),TAMSX3("DE_CONTA")[1]),"CT1_DESC01"))),; //DESCRIÇÃO CONTA CONTABIL,; //DESCRIÇÃO CONTA CONTABIL
					TiraGraf(Alltrim(POSICIONE("SB1",1,xFilial("SB1")+PADR(SD1->D1_COD,TAMSX3("D1_COD")[1]),"B1_DESC"))),; //PRODUTO,; //PRODUTO
					AllTrim(_cNaturez),; //NATUREZA
					(SD1->D1_TOTAL * SDE->DE_PERC)/100,; //VALOR
					SE2->E2_VALLIQ,; //VALOR PAGO
					SD1->D1_VALPIS,; //PIS
					SD1->D1_VALCOF,; //COFINS
					SD1->D1_VALCSL,; //CSLL
					SD1->D1_VALISS,; //ISS
					SD1->D1_VALIRR,; //IR
					SD1->D1_VALINS,; //INSS
					Alltrim(SD1->D1_DESCRI)}) //MENSAGEM 				
				
					SC7->(DbCloseArea())
				
				EndIf
		    Next _rr 
		    SD1->(DbCloseArea())
		    (_cAlias)->(DbCloseArea())				
		EndIf
		
	Case(_cAlias == "SEZ")     
		
		///////////////////////////////
		// Função para validar a ZZA //
		///////////////////////////////
		If(_lFlag)
			DbSelectArea("SE2")
			DbSetOrder(1)
			If(DbSeek(_cIndice)) 
				_cIndice := SE2->E2_FILIAL+SE2->E2_PREFIXO+SE2->E2_NUM+SE2->E2_PARCELA+SE2->E2_TIPO+SE2->E2_FORNECE+SE2->E2_LOJA	
			EndIf  
		EndIf	
		////////////////////////////////////////////////////////////
		
		DbSelectArea(_cAlias)
		DbSetOrder(1)
		If(DbSeek(_cIndice))       
		
			For _rr := 1 to Len(_xForn)
				
				If(SEZ->EZ_CLIFOR == _xForn[_rr])    
				
					DbSelectArea("SE2")
					DbSetOrder(1)
					DbSeek(xFilial("SE2")+SEZ->EZ_PREFIXO+SEZ->EZ_NUM+SEZ->EZ_PARCELA+SEZ->EZ_TIPO+SEZ->EZ_CLIFOR+SEZ->EZ_LOJA+SEZ->EZ_NATUREZ+SEZ->EZ_CCUSTO)
		
					Aadd(_cRDados,{SE2->E2_NUM,; //NUMERO DA NOTA FISCAL
					SE2->E2_PREFIXO,; //SERIE DA NOTA FISCAL
					SE2->E2_NUM,; //NUMERO DO PEDIDO DE COMPRAS
					SE2->E2_FORNECE,; //CODIGO DO FORNECEDOR
					TiraGraf(Alltrim(POSICIONE("SA2",1,xFilial("SA2")+PADR(SE2->E2_FORNECE,TAMSX3("E2_FORNECE")[1]),"A2_NOME"))),; //NOME DO FORNECEDOR
					SE2->E2_CONTAD,; //CONTA CONTABIL
					TiraGraf(Alltrim(POSICIONE("CT1",1,xFilial("CT1")+PADR(SE2->E2_CONTAD,TAMSX3("D1_CONTA")[1]),"CT1_DESC01"))),; //DESCRIÇÃO CONTA CONTABIL
					"",; //PRODUTO
					SE2->E2_NATUREZ,; //NATUREZA
					SE2->E2_VALOR,; //VALOR
					SE2->E2_VALLIQ,; //VALOR PAGO
					SE2->E2_PIS,; //PIS
					SE2->E2_COFINS,; //COFINS
					SE2->E2_CSLL,; //CSLL
					SE2->E2_ISS,; //ISS
					SE2->E2_IRRF,; //IR
					SE2->E2_INSS,; //INSS
					Alltrim(SE2->E2_HIST)}) //MENSAGEM 
				
				EndIf
				
			Next _rr
			SE2->(DbCloseArea())
		   	(_cAlias)->(DbCloseArea())
		EndIf		
		
		Case(_cAlias == "CV4") 
		
		///////////////////////////////
		// Função para validar a ZZA //
		///////////////////////////////
		If(_lFlag)
			_cAlias := "SE2"	
		EndIf	
		///////////////////////////////
		
		DbSelectArea(_cAlias)
		DbSetOrder(1)
		If(DbSeek(_cIndice))      
		
			If(!_lFlag)
				_cQuery3 := " SELECT E2_FILIAL, E2_PREFIXO, E2_NUM, E2_PARCELA, E2_TIPO, E2_FORNECE, E2_LOJA, E2_EMIS1"+ c_ent 
				_cQuery3 += " FROM "+RETSQLNAME("SE2")+" SE2"+ c_ent
				_cQuery3 += " INNER JOIN "+RETSQLNAME("CV4")+" CV4"+ c_ent
				_cQuery3 += " ON E2_FILIAL = CV4_FILIAL"+ c_ent
				_cQuery3 += " AND E2_EMIS1 = CV4_DTSEQ"+ c_ent
				_cQuery3 += " WHERE"+ c_ent
				_cQuery3 += " E2_ARQRAT LIKE '%"+_cIndice+"%'"+ c_ent
				_cQuery3 += " AND SE2.D_E_L_E_T_ = ' '"+ c_ent
				_cQuery3 += " AND CV4.D_E_L_E_T_ = ' '"+ c_ent 
				_cQuery3 += " GROUP BY E2_FILIAL, E2_PREFIXO, E2_NUM, E2_PARCELA, E2_TIPO, E2_FORNECE, E2_LOJA, E2_EMIS1"+ c_ent 
				_cQuery3 += " ORDER BY E2_EMIS1"+ c_ent
				_cQuery3 := ChangeQuery(_cQuery3)   
					
					If Select(_cTrab3) > 0
						dbSelectArea(_cTrab3)
						(_cTrab3)->(dbCloseArea())
					EndIf
					
					dbUseArea(.T.,"TOPCONN",TcGenQry(,,_cQuery3),_cTrab3,.T.,.T.)       
					
					dbSelectArea(_cTrab3)
					
					If(!Eof())
					  While !(_cTrab3)->(Eof())  
					    	_cIndice := (_cTrab3)->E2_FILIAL+(_cTrab3)->E2_PREFIXO+(_cTrab3)->E2_NUM+(_cTrab3)->E2_PARCELA+(_cTrab3)->E2_TIPO+(_cTrab3)->E2_FORNECE+(_cTrab3)-> E2_LOJA		
					  	(_cTrab3)->(DbSkip())
					  EndDo 
					EndIf  
					
					(_cTrab3)->(dbCloseArea()) 		    
			EndIf
			
			For _rr := 1 to Len(_xForn)          
			
			DbSelectArea("SE2")
			DbSetOrder(1)
			DbSeek(xFilial("SE2")+_cIndice)
			
				If(SE2->E2_FORNECE == _xForn[_rr])
		
					Aadd(_cRDados,{SE2->E2_NUM,; //NUMERO DA NOTA FISCAL
					SE2->E2_PREFIXO,; //SERIE DA NOTA FISCAL
					SE2->E2_NUM,; //NUMERO DO PEDIDO DE COMPRAS
					SE2->E2_FORNECE,; //CODIGO DO FORNECEDOR
					TiraGraf(Alltrim(POSICIONE("SA2",1,xFilial("SA2")+PADR(SE2->E2_FORNECE,TAMSX3("E2_FORNECE")[1]),"A2_NOME"))),; //NOME DO FORNECEDOR
					SE2->E2_CONTAD,; //CONTA CONTABIL
					TiraGraf(Alltrim(POSICIONE("CT1",1,xFilial("CT1")+PADR(SE2->E2_CONTAD,TAMSX3("D1_CONTA")[1]),"CT1_DESC01"))),; //DESCRIÇÃO CONTA CONTABIL
					"",; //PRODUTO
					SE2->E2_NATUREZ,; //NATUREZA
					(SE2->E2_VALOR * CV4->CV4_PERCEN) / 100,; //VALOR
					SE2->E2_VALLIQ,; //VALOR PAGO
					SE2->E2_PIS,; //PIS
					SE2->E2_COFINS,; //COFINS
					SE2->E2_CSLL,; //CSLL
					SE2->E2_ISS,; //ISS
					SE2->E2_IRRF,; //IR
					SE2->E2_INSS,; //INSS
					Alltrim(SE2->E2_HIST)}) //MENSAGEM 
				
				EndIf
				
			Next _rr 
			SE2->(DbCloseArea())
			(_cAlias)->(DbCloseArea())
		EndIf
		
EndCase

RestArea(_cArea5)

Return(_cRDados)
/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³OIFR004   ºAutor  ³Leandro Ribeiro     º Data ³  06/21/14   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Função para retornar o Mes de ingles para portugues.       º±±
±±º          ³                                                            º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP                                                        º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function ReturnMes(_zMes)
        
Local _cMes  := SUBSTR(CMonth(_zMes),1,3) 
Local _cRetM := ""

Do Case
	Case(_cMes == "Jan")
		_cRetM := "Jan"	
	Case(_cMes == "Feb")
		_cRetM := "Fev"	 
	Case(_cMes == "Mar")
		_cRetM := "Mar"	
	Case(_cMes == "Apr")
		_cRetM := "Abr"	 
	Case(_cMes == "May")
		_cRetM := "Mai"	
	Case(_cMes == "Jun")
		_cRetM := "Jun"	 
	Case(_cMes == "Jul")
		_cRetM := "Jul"	
	Case(_cMes == "Aug")
		_cRetM := "Ago"							
	Case(_cMes == "Sep")
		_cRetM := "Set"	
	Case(_cMes == "Oct")
		_cRetM := "Out"	 
	Case(_cMes == "Nov")
		_cRetM := "Nov"	
	Case(_cMes == "Dec")
		_cRetM := "Dez"	
EndCase		

Return(_cRetM)
                
/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³OIFR004   ºAutor  ³Leandro Ribeiro     º Data ³  28/05/14   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Função para Retirar os caracteres especias de uma string   º±±
±±º          ³                                                            º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP                                                        º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static function TiraGraf(_sOrig)

Local _sRet := _sOrig  

If(ValType(_sOrig) == "A")
	_sRet := _sOrig[1]
EndIf

If(!Empty(_sRet))
   
   _sRet := Strtran (_sRet, "á", "a")
   _sRet := Strtran (_sRet, "é", "e")
   _sRet := Strtran (_sRet, "í", "i")
   _sRet := Strtran (_sRet, "ó", "o")
   _sRet := Strtran (_sRet, "ú", "u")
   _sRet := Strtran (_sRet, "Á", "A")
   _sRet := Strtran (_sRet, "É", "E")
   _sRet := Strtran (_sRet, "Í", "I")
   _sRet := Strtran (_sRet, "Ó", "O")
   _sRet := Strtran (_sRet, "Ú", "U")
   _sRet := Strtran (_sRet, "ã", "a")
   _sRet := Strtran (_sRet, "õ", "o")
   _sRet := Strtran (_sRet, "Ã", "A")
   _sRet := Strtran (_sRet, "Õ", "O")
   _sRet := Strtran (_sRet, "â", "a")
   _sRet := Strtran (_sRet, "ê", "e")
   _sRet := Strtran (_sRet, "î", "i")
   _sRet := Strtran (_sRet, "ô", "o")
   _sRet := Strtran (_sRet, "û", "u")
   _sRet := Strtran (_sRet, "Â", "A")
   _sRet := Strtran (_sRet, "Ê", "E")
   _sRet := Strtran (_sRet, "Î", "I")
   _sRet := Strtran (_sRet, "Ô", "O")
   _sRet := Strtran (_sRet, "Û", "U")
   _sRet := Strtran (_sRet, "ç", "c")
   _sRet := Strtran (_sRet, "Ç", "C")
   _sRet := Strtran (_sRet, "à", "a")
   _sRet := Strtran (_sRet, "À", "A")
   _sRet := Strtran (_sRet, "º", " ")
   _sRet := Strtran (_sRet, "ª", " ")   
   _sRet := Strtran (_sRet, "´", " ")
   _sRet := Strtran (_sRet, "`", " ")
   _sRet := Strtran (_sRet, "^", " ")
   _sRet := Strtran (_sRet, "~", " ")
   _sRet := Strtran (_sRet, "!", " ")
   _sRet := Strtran (_sRet, "@", " ")
   _sRet := Strtran (_sRet, "#", " ")
   _sRet := Strtran (_sRet, "$", " ")
   _sRet := Strtran (_sRet, "%", " ")
   _sRet := Strtran (_sRet, "¨", " ")
   _sRet := Strtran (_sRet, "&", " ")
   _sRet := Strtran (_sRet, "*", " ")
   _sRet := Strtran (_sRet, "(", " ")
   _sRet := Strtran (_sRet, ")", " ")
   _sRet := Strtran (_sRet, "-", " ")
   _sRet := Strtran (_sRet, "_", " ")
   _sRet := Strtran (_sRet, "+", " ")
   _sRet := Strtran (_sRet, "=", " ")
   _sRet := Strtran (_sRet, "§", " ")
   _sRet := Strtran (_sRet, "¹", " ")
   _sRet := Strtran (_sRet, "²", " ")
   _sRet := Strtran (_sRet, "³", " ")
   _sRet := Strtran (_sRet, "£", " ")
   _sRet := Strtran (_sRet, "¢", " ")
   _sRet := Strtran (_sRet, "¬", " ")
   _sRet := Strtran (_sRet, "{", " ")
   _sRet := Strtran (_sRet, "[", " ")
   _sRet := Strtran (_sRet, "}", " ")
   _sRet := Strtran (_sRet, "}", " ")
   _sRet := Strtran (_sRet, ",", " ")
   _sRet := Strtran (_sRet, ".", " ")
   _sRet := Strtran (_sRet, ";", " ")
   _sRet := Strtran (_sRet, "/", " ")
   _sRet := Strtran (_sRet, "<", " ")
   _sRet := Strtran (_sRet, ">", " ")
   _sRet := Strtran (_sRet, ":", " ")
   _sRet := Strtran (_sRet, "?", " ")
   _sRet := Strtran (_sRet, "|", " ")
   _sRet := Strtran (_sRet, "\", " ")
   _sRet := Strtran (_sRet, "'", " ")
   _sRet := Strtran (_sRet, '"', " ")
   _sRet := Strtran (_sRet, chr (9), " ") // TAB 
   
EndIf
   
Return(_sRet)


