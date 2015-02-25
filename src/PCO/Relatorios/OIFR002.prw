#INCLUDE "Protheus.ch"
#INCLUDE "TopConn.ch"
#INCLUDE "rwmake.ch" 

#DEFINE c_ent CHR(13)+CHR(10)

#DEFINE _OPC_cGETFILE ( GETF_RETDIRECTORY + GETF_ONLYSERVER )


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³OIFR002   ºAutor  ³Leandro Ribeiro     º Data ³  13/06/14   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Programa para gerar planilha de conta com movimentações    º±±
±±º          ³ na tabela AKD de acordo com a parametrização do usuario.   º±±
±±º          ³ Planilha Conta Orçamentaria Realizado x Orçado  		      º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP                                                        º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

User Function OIFR002()  

Local _aArea := GetArea() 
Local cPerg  := Padr("OIFR002",Len(SX1->X1_GRUPO))      	             
                                                                                      //AL2A                                                                                                     ENG
//PutSx1(cPerg,"01","Tipo de Saldo"       , "", "", "mv_ch1" , "C", 08,0 ,0, "C", "", ""      ,"", "", "MV_PAR01", "PG" , "" , "" , "", "RE"      ,"","","EM","","","CT","","","","","",""/*,"","",""*///)     
PutSx1(cPerg,"01","Saldo 1"       		     , "", "", "mv_ch1" , "C", TAMSX3("AKD_TPSALD")[1],0 ,0, "C", "", ""    ,"", "", "MV_PAR01", "PR" , "" , "" , "" , "EM" , "" , "" , "RE" , "" , "" , "PG" , "","","PR+EM+RE","","","","","","","",""   ,"")
PutSx1(cPerg,"02","Saldo 2"                  , "", "", "mv_ch2" , "C", TAMSX3("AKD_TPSALD")[1],0 ,0, "C", "", ""    ,"", "", "MV_PAR02", "CT" , "" , "" , "" , "OR" , "" , "" , "OR+CT" , "" , "" , "" , "","","","","","","","","","",""   ,"")
PutSx1(cPerg,"03","Item Contabil de?"        , "", "", "mv_ch3" , "C", TAMSX3("AKD_ITCTB")[1] ,0 ,0, "G", "", "CTD" ,"", "", "MV_PAR03", "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"")
PutSx1(cPerg,"04","Item Contabil Ate?"       , "", "", "mv_ch4" , "C", TAMSX3("AKD_ITCTB")[1] ,0 ,0, "G", "", "CTD" ,"", "", "MV_PAR04", "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"")
PutSx1(cPerg,"05","Centro de Custo de?"      , "", "", "mv_ch5" , "C", TAMSX3("AKD_CC")[1]    ,0, 0, "G", "", "CTT" ,"", "", "MV_PAR05", "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"") 
PutSx1(cPerg,"06","Centro de Custo ate?"     , "", "", "mv_ch6" , "C", TAMSX3("AKD_CC")[1]    ,0, 0, "G", "", "CTT" ,"", "", "MV_PAR06", "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"") 
//PutSx1(cPerg,"07","Operação Orçamentaria" , "", "", "mv_ch7" , "C", 1                      ,0, 0, "C", "", " "   ,"", "", "MV_PAR07", "1" , "" , "" , "" , "2" , "" , "" , "3" , "" , "" , "" , "","","","","","","","","","",""   ,"")
PutSx1(cPerg,"07","Oper. Orçamentaria de?"   , "", "", "mv_ch7"  , "C", TAMSX3("AKF_CODIGO")[1]   ,0, 0, "G", "", "AKF"     ,"", "", "MV_PAR07" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"")
PutSx1(cPerg,"08","Oper. Orçamentaria Ate?"  , "", "", "mv_ch8"  , "C", TAMSX3("AKF_CODIGO")[1]   ,0, 0, "G", "", "AKF"     ,"", "", "MV_PAR08" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"")
PutSx1(cPerg,"09","Data"                     , "", "", "mv_ch9" , "D", TAMSX3("AKD_DATA")[1]  ,0, 0, "G", "", ""    ,"", "", "MV_PAR09", "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"")
//PutSx1(cPerg,"08","Data Ate"              , "", "", "mv_ch8" , "D", TAMSX3("AKD_DATA")[1]  ,0, 0, "G", "", ""        ,"", "", "MV_PAR08", "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"")
PutSx1(cPerg,"10","Conta Orçamentaria de?"   , "", "", "mv_ch10" , "C", TAMSX3("AK5_CODIGO")[1]   ,0, 0, "G", "", "AK5"     ,"", "", "MV_PAR10" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"") 
PutSx1(cPerg,"11","Conta Orçamentaria ate?"  , "", "", "mv_ch11" , "C", TAMSX3("AK5_CODIGO")[1]   ,0, 0, "G", "", "AK5"     ,"", "", "MV_PAR11" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"") 
       

If Pergunte(cPerg,.T.)
	
	Processa({|| GOIFR002() },"Aguarde...")
	
EndIf

RestArea(_aArea)
                                               
Return()


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³GOIFR002   ºAutor  ³Leandro Ribeiro    º Data ³  13/06/14   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Função para gerar a planilha de excel                      º±±
±±º          ³                                                            º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP                                                        º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function GOIFR002()

Local _cOIFR002  := GetArea()
//Local _cFile     := "OIFR002" + DToS( Date() ) + ".xml"                      
Local _cFile     := "OIFR002.xml"                      
Local _cPath     := GetTempPath() //Pega a Pasta de Arquivo temporário                 

Local _cQuery1   := "" 
Local _cTrab1  	 := GetNextAlias() 
Local _cTrab2  	 := GetNextAlias()
Local _cArray1   := {} //Retorna as Contas Orçamentarias Superior 
Local _cArray2	 := {} //Retorna as Contas Orçamentarias  
Local _cArray5   := {}
Local _cArray6   := {} 
Local _cArray7   := {}
Local _cArray8   := {}
Local _cArray9   := {}
Local _cVerSald  := "" 
Local _cFlag1    := .F. 
Local _cCont     := 0
Local _cRetSoma  := ""  
Local _cRetMes	 := ReturnMes(MV_PAR09)    

Private _cUltimoD  := SUBSTR(DTOS(MV_PAR09),1,6)+AllTrim(STR(Last_Day(SUBSTR(DTOS(MV_PAR09),1,6)+"01")))
Private _cFirstD   := SUBSTR(DTOS(MV_PAR09),1,6)+"01"                   
Private _cMV_PAR01 := ""
Private _cMV_PAR02 := ""
Private cPath      := "" 
Private _nHandle 
                      
dbSelectArea("AK5")
dbSetOrder(1)      
dbSelectArea("AKD")
dbSetOrder(1)           

Do Case
	Case(MV_PAR01 == 1)
		_cMV_PAR01 := "('PR')"
	Case(MV_PAR01 == 2)
		_cMV_PAR01 := "('EM')"			
	Case(MV_PAR01 == 3)
		_cMV_PAR01 := "('RE')"
	Case(MV_PAR01 == 4)
		_cMV_PAR01 := "('PG')"
	Case(MV_PAR01 == 5)
		_cMV_PAR01 := "('PR','EM','RE')"
EndCase	 
Do Case
	Case(MV_PAR02 == 1)
		_cMV_PAR02 := "('CT')"
	Case(MV_PAR02 == 2)
		_cMV_PAR02 := "('OR')"			
	Case(MV_PAR02 == 3)
		_cMV_PAR02 := "('OR','CT')"
EndCase		
	
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

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ//
//"Saldo Realizado"			MV_PAR01  //
//"Saldo Orçado"			MV_PAR02  //
//"Item Contabil de?" 		MV_PAR03  //
//"Item Contabil Ate?" 		MV_PAR04  //
//"Centro de Custo de?" 	MV_PAR05  //
//"Centro de Custo ate?" 	MV_PAR06  //
//"Operação Orçamentaria" 	MV_PAR07  //
//"Data" 			   		MV_PAR08  //
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ//     

_cQuery1 := " SELECT DISTINCT AK5_COSUP " + c_ent
_cQuery1 += " FROM "+RETSQLNAME("AKD")+" AKD "+ c_ent
_cQuery1 += " INNER JOIN "+RETSQLNAME("AK5")+" AK5 "+ c_ent
_cQuery1 += " ON AKD_FILIAL = AK5_FILIAL "+ c_ent 
_cQuery1 += " AND AKD_CO = AK5_CODIGO "+ c_ent
_cQuery1 += " WHERE "+ c_ent
//_cQuery1 += " AKD_TPSALD IN "+AllTrim(_cMV_PAR01)+" "+ c_ent
//_cQuery1 += " AND AKD_CC BETWEEN '"+AllTrim(MV_PAR05)+"' AND '"+AllTrim(MV_PAR06)+"' "+ c_ent
_cQuery1 += " AKD_CC BETWEEN '"+AllTrim(MV_PAR05)+"' AND '"+AllTrim(MV_PAR06)+"' "+ c_ent
_cQuery1 += " AND AKD_ITCTB BETWEEN '"+AllTrim(MV_PAR03)+"' AND '"+AllTrim(MV_PAR04)+"' "+ c_ent
//_cQuery1 += " AND AKD_OPER = '"+Alltrim(MV_PAR07)+"' "+ c_ent
_cQuery1 += " AND AKD_OPER BETWEEN '"+AllTrim(MV_PAR07)+"' AND '"+AllTrim(MV_PAR08)+"' "+ c_ent
_cQuery1 += " AND AKD_CO BETWEEN '"+MV_PAR10+"' AND '"+MV_PAR11+"' "+ c_ent  
_cQuery1 += " AND AKD_DATA BETWEEN '"+_cFirstD+"' AND '"+_cUltimoD+"' "+ c_ent
_cQuery1 += " AND AKD.D_E_L_E_T_ = ' ' "+ c_ent
_cQuery1 += " AND AK5.D_E_L_E_T_ = ' ' "+ c_ent
_cQuery1 := ChangeQuery(_cQuery1) 


If Select(_cTrab1) > 0
	dbSelectArea(_cTrab1)
	(_cTrab1)->(dbCloseArea())
EndIf

dbUseArea(.T.,"TOPCONN",TcGenQry(,,_cQuery1),_cTrab1,.T.,.T.)       

dbSelectArea(_cTrab1)

If(!Eof())
  While !(_cTrab1)->(Eof())  
		Aadd(_cArray1,AllTrim((_cTrab1)->AK5_COSUP))
  	DbSkip()
  EndDo 
EndIf  

(_cTrab1)->(dbCloseArea())    


/*_cQuery1 := " SELECT DISTINCT AK5_COSUP " + c_ent
_cQuery1 += " FROM "+RETSQLNAME("AKD")+" AKD "+ c_ent
_cQuery1 += " INNER JOIN "+RETSQLNAME("AK5")+" AK5 "+ c_ent
_cQuery1 += " ON AKD_FILIAL = AK5_FILIAL "+ c_ent 
_cQuery1 += " AND AKD_CO = AK5_CODIGO "+ c_ent
_cQuery1 += " WHERE "+ c_ent
_cQuery1 += " AKD_TPSALD IN "+AllTrim(_cMV_PAR02)+" "+ c_ent
_cQuery1 += " AND AKD_CC BETWEEN '"+AllTrim(MV_PAR05)+"' AND '"+AllTrim(MV_PAR06)+"' "+ c_ent
_cQuery1 += " AND AKD_ITCTB BETWEEN '"+AllTrim(MV_PAR03)+"' AND '"+AllTrim(MV_PAR04)+"' "+ c_ent
//_cQuery1 += " AND AKD_OPER = '"+Alltrim(MV_PAR07)+"' "+ c_ent
_cQuery1 += " AND AKD_OPER BETWEEN '"+AllTrim(MV_PAR07)+"' AND '"+AllTrim(MV_PAR08)+"' "+ c_ent
_cQuery1 += " AND AKD_CO BETWEEN '"+MV_PAR10+"' AND '"+MV_PAR11+"' "+ c_ent  
_cQuery1 += " AND AKD_DATA BETWEEN '"+_cFirstD+"' AND '"+_cUltimoD+"' "+ c_ent
_cQuery1 += " AND AKD.D_E_L_E_T_ = ' ' "+ c_ent
_cQuery1 += " AND AK5.D_E_L_E_T_ = ' ' "+ c_ent
_cQuery1 := ChangeQuery(_cQuery1) 


If Select(_cTrab1) > 0
	dbSelectArea(_cTrab1)
	(_cTrab1)->(dbCloseArea())
EndIf

dbUseArea(.T.,"TOPCONN",TcGenQry(,,_cQuery1),_cTrab1,.T.,.T.)       

dbSelectArea(_cTrab1)

If(!Eof())
  While !(_cTrab1)->(Eof())  
		Aadd(_cArray1,AllTrim((_cTrab1)->AK5_COSUP))
  	DbSkip()
  EndDo 
EndIf  

(_cTrab1)->(dbCloseArea())*/ 

If(EMPTY(_cArray1))
   Aviso("Aviso","Nenhum dado encontrado com os parametros passados!",{"OK"}) 
   		U_OIFR002()
   Return()
EndIf
            
fWrite(_nHandle,'<?xml version="1.0"?>'+ c_ent)
fWrite(_nHandle,'<?mso-application progid="Excel.Sheet"?>'+ c_ent)
fWrite(_nHandle,'<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"'+ c_ent)
fWrite(_nHandle,' xmlns:o="urn:schemas-microsoft-com:office:office"'+ c_ent)
fWrite(_nHandle,' xmlns:x="urn:schemas-microsoft-com:office:excel"'+ c_ent)
fWrite(_nHandle,' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"'+ c_ent)
fWrite(_nHandle,' xmlns:html="http://www.w3.org/TR/REC-html40">'+ c_ent)
fWrite(_nHandle,' <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">'+ c_ent)
fWrite(_nHandle,'  <Author>PROTHEUS 11</Author>'+ c_ent)
fWrite(_nHandle,'  <LastAuthor>PROTHEUS 11</LastAuthor>'+ c_ent)
fWrite(_nHandle,'  <Created>2013-05-23T13:58:13Z</Created>'+ c_ent)
fWrite(_nHandle,'  <LastSaved>2014-06-13T18:22:26Z</LastSaved>'+ c_ent)
fWrite(_nHandle,'  <Company>Microsoft</Company>'+ c_ent)
fWrite(_nHandle,'  <Version>12.00</Version>'+ c_ent)
fWrite(_nHandle,' </DocumentProperties>'+ c_ent)
fWrite(_nHandle,' <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">'+ c_ent)
fWrite(_nHandle,'  <WindowHeight>7050</WindowHeight>'+ c_ent)
fWrite(_nHandle,'  <WindowWidth>11955</WindowWidth>'+ c_ent)
fWrite(_nHandle,'  <WindowTopX>240</WindowTopX>'+ c_ent)
fWrite(_nHandle,'  <WindowTopY>360</WindowTopY>'+ c_ent)
fWrite(_nHandle,'  <TabRatio>921</TabRatio>'+ c_ent)
fWrite(_nHandle,'  <ProtectStructure>False</ProtectStructure>'+ c_ent)
fWrite(_nHandle,'  <ProtectWindows>False</ProtectWindows>'+ c_ent)
fWrite(_nHandle,' </ExcelWorkbook>'+ c_ent)
fWrite(_nHandle,' <Styles>'+ c_ent)
fWrite(_nHandle,'  <Style ss:ID="Default" ss:Name="Normal">'+ c_ent)
fWrite(_nHandle,'   <Alignment ss:Vertical="Bottom"/>'+ c_ent)
fWrite(_nHandle,'   <Borders/>'+ c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>'+ c_ent)
fWrite(_nHandle,'   <Interior/>'+ c_ent)
fWrite(_nHandle,'   <NumberFormat/>'+ c_ent)
fWrite(_nHandle,'   <Protection/>'+ c_ent)
fWrite(_nHandle,'  </Style>'+ c_ent)
fWrite(_nHandle,'  <Style ss:ID="s62">'+ c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>'+ c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>'+ c_ent)
fWrite(_nHandle,'  </Style>'+ c_ent)
fWrite(_nHandle,'  <Style ss:ID="s63">'+ c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Color="#000000"/>'+ c_ent)
fWrite(_nHandle,'  </Style>'+ c_ent)
fWrite(_nHandle,'  <Style ss:ID="s64">'+ c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>'+ c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Color="#000000"/>'+ c_ent)
fWrite(_nHandle,'  </Style>'+ c_ent)
fWrite(_nHandle,'  <Style ss:ID="s65">'+ c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom" ss:Indent="8"/>'+ c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Color="#000000" ss:Bold="1"/>'+ c_ent)
fWrite(_nHandle,'  </Style>'+ c_ent)
fWrite(_nHandle,'  <Style ss:ID="s66">'+ c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>'+ c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Color="#000000"/>'+ c_ent)
fWrite(_nHandle,'  </Style>'+ c_ent)
fWrite(_nHandle,'  <Style ss:ID="s67">'+ c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom" ss:Indent="8"/>'+ c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Color="#000000"/>'+ c_ent)
fWrite(_nHandle,'  </Style>'+ c_ent)
fWrite(_nHandle,'  <Style ss:ID="s68">'+ c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>'+ c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Color="#FFFFFF" ss:Bold="1"/>'+ c_ent)
fWrite(_nHandle,'   <Interior ss:Color="#7F7F7F" ss:Pattern="Solid"/>'+ c_ent)
fWrite(_nHandle,'  </Style>'+ c_ent)
fWrite(_nHandle,'  <Style ss:ID="s69">'+ c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>'+ c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Color="#FFFFFF"/>'+ c_ent)
fWrite(_nHandle,'   <Interior ss:Color="#7F7F7F" ss:Pattern="Solid"/>'+ c_ent)
fWrite(_nHandle,'  </Style>'+ c_ent)
fWrite(_nHandle,'  <Style ss:ID="s71">'+ c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>'+ c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Color="#000000"/>'+ c_ent)
fWrite(_nHandle,'   <Interior/>'+ c_ent)
fWrite(_nHandle,'  </Style>'+ c_ent)
fWrite(_nHandle,'  <Style ss:ID="s72">'+ c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>'+ c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Color="#000000"/>'+ c_ent)
fWrite(_nHandle,'   <Interior/>'+ c_ent)
fWrite(_nHandle,'  </Style>'+ c_ent)
fWrite(_nHandle,'  <Style ss:ID="s73">'+ c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>'+ c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Color="#000000" ss:Bold="1"/>'+ c_ent)
fWrite(_nHandle,'   <Interior ss:Color="#D8D8D8" ss:Pattern="Solid"/>'+ c_ent)
fWrite(_nHandle,'  </Style>'+ c_ent)   
fWrite(_nHandle,'  <Style ss:ID="s74">'+ c_ent) 
fWrite(_nHandle,'     <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>'+ c_ent) 
fWrite(_nHandle,'     <Font ss:FontName="Calibri" x:Family="Swiss" ss:Color="#993300"/>'+ c_ent) 
fWrite(_nHandle,'     <Interior ss:Color="#D8D8D8" ss:Pattern="Solid"/>'+ c_ent) 
fWrite(_nHandle,'     <NumberFormat ss:Format="Fixed"/>'+ c_ent) 
fWrite(_nHandle,'    </Style>'+ c_ent) 
fWrite(_nHandle,'    <Style ss:ID="s75">'+ c_ent) 
fWrite(_nHandle,'     <Alignment ss:Horizontal="Left" ss:Vertical="Bottom" ss:Indent="2"/>'+ c_ent) 
fWrite(_nHandle,'     <Font ss:FontName="Calibri" x:Family="Swiss" ss:Color="#000000"/>'+ c_ent) 
fWrite(_nHandle,'     <Interior/>'+ c_ent) 
fWrite(_nHandle,'    </Style>'+ c_ent) 
fWrite(_nHandle,'    <Style ss:ID="s76">'+ c_ent) 
fWrite(_nHandle,'     <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>'+ c_ent) 
fWrite(_nHandle,'     <Font ss:FontName="Calibri" x:Family="Swiss" ss:Color="#000000"/>'+ c_ent) 
fWrite(_nHandle,'     <Interior/>'+ c_ent) 
fWrite(_nHandle,'     <NumberFormat ss:Format="Fixed"/>'+ c_ent) 
fWrite(_nHandle,'    </Style>'+ c_ent) 
fWrite(_nHandle,'    <Style ss:ID="s77">'+ c_ent) 
fWrite(_nHandle,'     <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>'+ c_ent) 
fWrite(_nHandle,'     <Font ss:FontName="Calibri" x:Family="Swiss" ss:Color="#FFFFFF" ss:Bold="1"/>'+ c_ent) 
fWrite(_nHandle,'     <Interior ss:Color="#7F7F7F" ss:Pattern="Solid"/>'+ c_ent) 
fWrite(_nHandle,'    </Style>'+ c_ent) 
fWrite(_nHandle,'    <Style ss:ID="s78">'+ c_ent) 
fWrite(_nHandle,'     <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>'+ c_ent) 
fWrite(_nHandle,'     <Font ss:FontName="Calibri" x:Family="Swiss" ss:Color="#FFFFFF"/>'+ c_ent) 
fWrite(_nHandle,'     <Interior ss:Color="#7F7F7F" ss:Pattern="Solid"/>'+ c_ent) 
fWrite(_nHandle,'     <NumberFormat ss:Format="Fixed"/>'+ c_ent) 
fWrite(_nHandle,'    </Style>'+ c_ent) 
fWrite(_nHandle,'  <Style ss:ID="s79">'+ c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>'+ c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Color="#000000"/>'+ c_ent)
fWrite(_nHandle,'   <Interior/>'+ c_ent)
fWrite(_nHandle,'   <NumberFormat ss:Format="Fixed"/>'+ c_ent)
fWrite(_nHandle,'  </Style>'+ c_ent)
fWrite(_nHandle,'  <Style ss:ID="s80">'+ c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>'+ c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Color="#FFFFFF"/>'+ c_ent)
fWrite(_nHandle,'   <Interior ss:Color="#7F7F7F" ss:Pattern="Solid"/>'+ c_ent)
fWrite(_nHandle,'   <NumberFormat ss:Format="Fixed"/>'+ c_ent)
fWrite(_nHandle,'  </Style>'+ c_ent)
fWrite(_nHandle,' </Styles>'+ c_ent)
fWrite(_nHandle,' <Worksheet ss:Name="conta mes x orcto">'+ c_ent)
fWrite(_nHandle,'  <Table ss:ExpandedColumnCount="16" ss:ExpandedRowCount="10000" x:FullColumns="1"'+ c_ent)
fWrite(_nHandle,'   x:FullRows="1" ss:DefaultRowHeight="15">'+ c_ent)
fWrite(_nHandle,'   <Column ss:AutoFitWidth="0" ss:Width="15.75"/>'+ c_ent)
//fWrite(_nHandle,'   <Column ss:AutoFitWidth="0" ss:Width="165"/>'+ c_ent)
fWrite(_nHandle,'   <Column ss:Width="238.5"/>'+ c_ent)
fWrite(_nHandle,'   <Column ss:AutoFitWidth="0" ss:Width="75.75"/>'+ c_ent)
fWrite(_nHandle,'   <Column ss:StyleID="s62" ss:AutoFitWidth="0" ss:Width="48.75"/>'+ c_ent)
fWrite(_nHandle,'   <Column ss:StyleID="s62" ss:AutoFitWidth="0" ss:Width="57"/>'+ c_ent)
fWrite(_nHandle,'   <Column ss:StyleID="s62" ss:AutoFitWidth="0" ss:Width="60.75"/>'+ c_ent)
fWrite(_nHandle,'   <Column ss:Index="11" ss:Width="78.75"/>'+ c_ent)
fWrite(_nHandle,'   <Row ss:AutoFitHeight="0">'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s63"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s63"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s63"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'   </Row>'+ c_ent)
fWrite(_nHandle,'   <Row ss:AutoFitHeight="0">'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s63"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s65"><Data ss:Type="String">Oi Futuro</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s65"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s66"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s66"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'   </Row>'+ c_ent)
fWrite(_nHandle,'   <Row ss:AutoFitHeight="0">'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s63"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s65"><Data ss:Type="String">Diretoria de Planejamento e Desempenho</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s65"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s66"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s66"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s66"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'   </Row>'+ c_ent)
fWrite(_nHandle,'   <Row ss:AutoFitHeight="0">'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s63"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s65"><Data ss:Type="String">Estrutura Oi Futuro</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s65"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s66"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s66"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s66"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'   </Row>'+ c_ent)
fWrite(_nHandle,'   <Row ss:AutoFitHeight="0">'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s63"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s67"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s67"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s66"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s66"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s66"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'   </Row>'+ c_ent)
fWrite(_nHandle,'   <Row ss:AutoFitHeight="0">'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s63"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s67"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s67"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'   </Row>'+ c_ent)
fWrite(_nHandle,'   <Row ss:AutoFitHeight="0">'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s63"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s68"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s68"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s69"><Data ss:Type="String">Realizado</Data><Comment'+ c_ent)
fWrite(_nHandle,'      ss:Author="Tatiana Do Nascimento"><ss:Data'+ c_ent)
fWrite(_nHandle,'       xmlns="http://www.w3.org/TR/REC-html40"><B><Font html:Face="Tahoma"'+ c_ent)
fWrite(_nHandle,'         x:Family="Swiss" html:Size="9" html:Color="#000000">selecionar o mes</Font></B></ss:Data></Comment></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s69"><Data ss:Type="String">Orcamento</Data><Comment'+ c_ent)
fWrite(_nHandle,'      ss:Author="Tatiana Do Nascimento"><ss:Data'+ c_ent)
fWrite(_nHandle,'       xmlns="http://www.w3.org/TR/REC-html40"><B><Font html:Face="Tahoma"'+ c_ent)
fWrite(_nHandle,'         x:Family="Swiss" html:Size="9" html:Color="#000000">selecionar o mes</Font></B></ss:Data></Comment></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:MergeDown="1" ss:StyleID="s69"><Data ss:Type="String">Variacao Real x Orcto</Data></Cell>'+ c_ent)
fWrite(_nHandle,'   </Row>'+ c_ent)
fWrite(_nHandle,'   <Row ss:AutoFitHeight="0">'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s63"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s68"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s68"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s69"><Data ss:Type="String">'+_cRetMes+'</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s69"><Data ss:Type="String">'+_cRetMes+'</Data></Cell>'+ c_ent)
fWrite(_nHandle,'   </Row>'+ c_ent)
fWrite(_nHandle,'   <Row ss:AutoFitHeight="0" ss:Height="4.5">'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s63"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s72"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s72"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s72"/>'+ c_ent)
fWrite(_nHandle,'   </Row>'+ c_ent)
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/*For _xx := 1 to Len(_cArray1)    
	
	_cArray6 := RETCONTA(_cArray1[_xx]) 
	
	_cFlag1 := .F.    
	
		For _rr := 1 to Len(_cArray6)
	
			_cVerSald := VERISALDO(_cArray6[_rr][1],_cArray6[_rr][2]) 
	
				If!(AllTrim(_cVerSald) == "0" .OR. EMPTY(AllTrim(_cVerSald)))
	
					Aadd(_cArray7,_cArray6[_rr][1],_cArray6[_rr][2])	
	
					_cFlag1 := .T.
	
				EndIf
	
		Next _rr
		
		If(_cFlag1)
			Aadd(_cArray8,_cArray1[_xx])
		EndIf
Next _xx
            
_cArray1 := {}
_cArray2 := {}
_cArray1 := _cArray8
_cArray2 := _cArray7  */


For _cc := 1 to Len(_cArray1)         

		_cArray2 := RETCONTA(_cArray1[_cc])
	
		fWrite(_nHandle,'   <Row ss:AutoFitHeight="0">'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s63"/>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s73"><Data ss:Type="String">'+TiraGraf(Alltrim(POSICIONE("AK5",1,xFilial("AK5")+PADR(_cArray1[_cc],TAMSX3("AK5_CODIGO")[1]),"AK5_DESCRI")))+'</Data></Cell>'+ c_ent)
//		fWrite(_nHandle,'    <Cell ss:StyleID="s73"/>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s73"><Data ss:Type="String">'+Alltrim(_cArray1[_cc])+'</Data></Cell>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s74" ss:Formula="=SUM(R[1]C:R['+Alltrim(STR(Len(_cArray2)))+']C)"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s74" ss:Formula="=SUM(R[1]C:R['+Alltrim(STR(Len(_cArray2)))+']C)"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s74" ss:Formula="=SUM(R[1]C:R['+Alltrim(STR(Len(_cArray2)))+']C)"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
		fWrite(_nHandle,'   </Row>'+ c_ent)                                                                                 
		
		_cCont := _cCont + 2
		
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// CONTAS ORÇAMENTARIAS
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////         
		
			For _zz := 1 to Len(_cArray2)
				
				Aadd(_cArray5,AllTrim(CONTREAL(_cArray2[_zz][1])))
				Aadd(_cArray5,AllTrim(CONTORC(_cArray2[_zz][1])))
				Aadd(_cArray5,AllTrim(CONTSALD(_cArray2[_zz][1])))
			
			
				fWrite(_nHandle,'   <Row ss:AutoFitHeight="0">'+ c_ent)
				fWrite(_nHandle,'    <Cell ss:StyleID="s63"/>'+ c_ent)
				fWrite(_nHandle,'    <Cell ss:StyleID="s75"><Data ss:Type="String">'+TiraGraf(Alltrim(POSICIONE("AK5",1,xFilial("AK5")+PADR(_cArray2[_zz][1],TAMSX3("AK5_CODIGO")[1]),"AK5_DESCRI")))+'</Data></Cell>'+ c_ent)
//				fWrite(_nHandle,'    <Cell ss:StyleID="s75"/>'+ c_ent)
				fWrite(_nHandle,'    <Cell ss:StyleID="s75"><Data ss:Type="String">'+Alltrim(_cArray2[_zz][1])+'</Data></Cell>'+ c_ent)
				fWrite(_nHandle,'    <Cell ss:StyleID="s76"><Data ss:Type="Number">'+Strtran(Alltrim(Transform(Val(_cArray5[1]), "@E 999999999.99" )),",",".")+'</Data></Cell>'+ c_ent)
				fWrite(_nHandle,'    <Cell ss:StyleID="s76"><Data ss:Type="Number">'+Strtran(Alltrim(Transform(Val(_cArray5[2]), "@E 999999999.99" )),",",".")+'</Data></Cell>'+ c_ent)
				fWrite(_nHandle,'    <Cell ss:StyleID="s76"><Data ss:Type="Number">'+Strtran(Alltrim(Transform(Val(_cArray5[3]), "@E 999999999.99" )),",",".")+'</Data></Cell>'+ c_ent)
				fWrite(_nHandle,'   </Row>'+ c_ent)		
				
				_cArray5 := {}
				
				_cCont := _cCont + 1
				
			Next _zz  
			
			Aadd(_cArray9,_cCont)  

			_cCont := 0
		
		fWrite(_nHandle,'   <Row ss:AutoFitHeight="0">'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s63"/>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s75"/>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s75"/>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s72"/>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s72"/>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s72"/>'+ c_ent)
		fWrite(_nHandle,'   </Row>'+ c_ent)                       
		
Next _cc

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////		

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

_cRetSoma := LINHACALC(_cArray9)      

fWrite(_nHandle,'   <Row ss:AutoFitHeight="0">'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s63"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s77"><Data ss:Type="String">TOTAL GERAL</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s77"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s78" ss:Formula="='+_cRetSoma+'"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s78" ss:Formula="='+_cRetSoma+'"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s78" ss:Formula="='+_cRetSoma+'"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
fWrite(_nHandle,'   </Row>'+ c_ent)
fWrite(_nHandle,'   <Row ss:AutoFitHeight="0">'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s63"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s75"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s75"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s72"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s72"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s72"/>'+ c_ent)
fWrite(_nHandle,'   </Row>'+ c_ent)
//fWrite(_nHandle,'   <Row ss:Index="21" ss:AutoFitHeight="0">'+ c_ent)
fWrite(_nHandle,'   <Row ss:AutoFitHeight="0"> '+ c_ent)
fWrite(_nHandle,'    <Cell ss:Index="2" ss:StyleID="s71"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
fWrite(_nHandle,'   </Row>'+ c_ent)
fWrite(_nHandle,'  </Table>'+ c_ent)
fWrite(_nHandle,'  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">'+ c_ent)
fWrite(_nHandle,'   <PageSetup>'+ c_ent)
fWrite(_nHandle,'    <Header x:Margin="0.31496062000000002"/>'+ c_ent)
fWrite(_nHandle,'    <Footer x:Margin="0.31496062000000002"/>'+ c_ent)
fWrite(_nHandle,'    <PageMargins x:Bottom="0.78740157499999996" x:Left="0.511811024"'+ c_ent)
fWrite(_nHandle,'     x:Right="0.511811024" x:Top="0.78740157499999996"/>'+ c_ent)
fWrite(_nHandle,'   </PageSetup>'+ c_ent)
fWrite(_nHandle,'   <Unsynced/>'+ c_ent)
fWrite(_nHandle,'   <NoSummaryRowsBelowDetail/>'+ c_ent)
fWrite(_nHandle,'   <Print>'+ c_ent)
fWrite(_nHandle,'    <ValidPrinterInfo/>'+ c_ent)
fWrite(_nHandle,'    <PaperSizeIndex>9</PaperSizeIndex>'+ c_ent)
fWrite(_nHandle,'    <HorizontalResolution>600</HorizontalResolution>'+ c_ent)
fWrite(_nHandle,'    <VerticalResolution>600</VerticalResolution>'+ c_ent)
fWrite(_nHandle,'   </Print>'+ c_ent)
fWrite(_nHandle,'   <TabColorIndex>63</TabColorIndex>'+ c_ent)
fWrite(_nHandle,'   <Selected/>'+ c_ent)
fWrite(_nHandle,'   <DoNotDisplayGridlines/>'+ c_ent)
fWrite(_nHandle,'   <Panes>'+ c_ent)
fWrite(_nHandle,'    <Pane>'+ c_ent)
fWrite(_nHandle,'     <Number>3</Number>'+ c_ent)
fWrite(_nHandle,'     <ActiveRow>12</ActiveRow>'+ c_ent)
fWrite(_nHandle,'     <ActiveCol>5</ActiveCol>'+ c_ent)
fWrite(_nHandle,'    </Pane>'+ c_ent)
fWrite(_nHandle,'   </Panes>'+ c_ent)
fWrite(_nHandle,'   <ProtectObjects>False</ProtectObjects>'+ c_ent)
fWrite(_nHandle,'   <ProtectScenarios>False</ProtectScenarios>'+ c_ent)
fWrite(_nHandle,'  </WorksheetOptions>'+ c_ent)
fWrite(_nHandle,' </Worksheet>'+ c_ent)
fWrite(_nHandle,'</Workbook>'+ c_ent)

//****************|
//Fechando arquivo|
//****************|
fClose(_nHandle)

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³Abre a Planilha Excel³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
ShellExecute( "Open", _cArquivo, " ", _cPath, 3 )

RestArea(_cOIFR002)

Return()   

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³OIFR002   ºAutor  ³Leandro Ribeiro     º Data ³  19/06/14   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Função para retornar contas atraves das contas superiores  º±±
±±º          ³                                                            º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP                                                        º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
                                  
Static Function RETCONTA(_cContSup)
                    
Local _cArea3  := GetArea()
Local _cQuery3 := "" 
Local _cTrab3  := GetNextAlias() 
Local _cArray3 := {}        

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ//
//"Saldo Realizado"			MV_PAR01  //
//"Saldo Orçado"			MV_PAR02  //
//"Item Contabil de?" 		MV_PAR03  //
//"Item Contabil Ate?" 		MV_PAR04  //
//"Centro de Custo de?" 	MV_PAR05  //
//"Centro de Custo ate?" 	MV_PAR06  //
//"Operação Orçamentaria" 	MV_PAR07  //
//"Data" 			   		MV_PAR08  //
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ//    

//_cQuery3 := " SELECT DISTINCT AKD_CO, AKD_CC " + c_ent
_cQuery3 := " SELECT DISTINCT AKD_CO " + c_ent
_cQuery3 += " FROM "+RETSQLNAME("AKD")+" AKD "+ c_ent
_cQuery3 += " INNER JOIN "+RETSQLNAME("AK5")+" AK5 "+ c_ent
_cQuery3 += " ON AKD_FILIAL = AK5_FILIAL "+ c_ent 
_cQuery3 += " AND AKD_CO = AK5_CODIGO "+ c_ent
_cQuery3 += " WHERE "+ c_ent
//_cQuery3 += " AKD_TPSALD IN "+AllTrim(_cMV_PAR01)+" "+ c_ent
//_cQuery3 += " AND AKD_CC BETWEEN '"+AllTrim(MV_PAR05)+"' AND '"+AllTrim(MV_PAR06)+"' "+ c_ent
_cQuery3 += " AKD_CC BETWEEN '"+AllTrim(MV_PAR05)+"' AND '"+AllTrim(MV_PAR06)+"' "+ c_ent
_cQuery3 += " AND AKD_ITCTB BETWEEN '"+AllTrim(MV_PAR03)+"' AND '"+AllTrim(MV_PAR04)+"' "+ c_ent
//_cQuery3 += " AND AKD_OPER = '"+Alltrim(MV_PAR07)+"' "+ c_ent   
_cQuery3 += " AND AKD_OPER BETWEEN '"+AllTrim(MV_PAR07)+"' AND '"+AllTrim(MV_PAR08)+"' "+ c_ent
_cQuery3 += " AND AKD_CO BETWEEN '"+MV_PAR10+"' AND '"+MV_PAR11+"' "+ c_ent  
_cQuery3 += " AND AK5_COSUP = '"+_cContSup+"' "+ c_ent   
_cQuery3 += " AND AKD_DATA BETWEEN '"+_cFirstD+"' AND '"+_cUltimoD+"' "+ c_ent
_cQuery3 += " AND AKD.D_E_L_E_T_ = ' ' "+ c_ent
_cQuery3 += " AND AK5.D_E_L_E_T_ = ' ' "+ c_ent
_cQuery3 := ChangeQuery(_cQuery3) 


If Select(_cTrab3) > 0
	dbSelectArea(_cTrab3)
	(_cTrab3)->(dbCloseArea())
EndIf

dbUseArea(.T.,"TOPCONN",TcGenQry(,,_cQuery3),_cTrab3,.T.,.T.)       

dbSelectArea(_cTrab3)

If(!Eof())
  While !(_cTrab3)->(Eof())  
		//Aadd(_cArray3,{AllTrim((_cTrab3)->AKD_CO),AllTrim((_cTrab3)->AKD_CC)})
		Aadd(_cArray3,{AllTrim((_cTrab3)->AKD_CO)})
  	DbSkip()
  EndDo 
EndIf  

/*(_cTrab3)->(dbCloseArea())  

//_cQuery3 := " SELECT DISTINCT AKD_CO, AKD_CC " + c_ent
_cQuery3 := " SELECT DISTINCT AKD_CO " + c_ent
_cQuery3 += " FROM "+RETSQLNAME("AKD")+" AKD "+ c_ent
_cQuery3 += " INNER JOIN "+RETSQLNAME("AK5")+" AK5 "+ c_ent
_cQuery3 += " ON AKD_FILIAL = AK5_FILIAL "+ c_ent 
_cQuery3 += " AND AKD_CO = AK5_CODIGO "+ c_ent
_cQuery3 += " WHERE "+ c_ent
_cQuery3 += " AKD_TPSALD IN "+AllTrim(_cMV_PAR02)+" "+ c_ent
_cQuery3 += " AND AKD_CC BETWEEN '"+AllTrim(MV_PAR05)+"' AND '"+AllTrim(MV_PAR06)+"' "+ c_ent
_cQuery3 += " AND AKD_ITCTB BETWEEN '"+AllTrim(MV_PAR03)+"' AND '"+AllTrim(MV_PAR04)+"' "+ c_ent
//_cQuery3 += " AND AKD_OPER = '"+Alltrim(MV_PAR07)+"' "+ c_ent   
_cQuery3 += " AND AKD_OPER BETWEEN '"+AllTrim(MV_PAR07)+"' AND '"+AllTrim(MV_PAR08)+"' "+ c_ent
_cQuery3 += " AND AKD_CO BETWEEN '"+MV_PAR10+"' AND '"+MV_PAR11+"' "+ c_ent  
_cQuery3 += " AND AK5_COSUP = '"+_cContSup+"' "+ c_ent   
_cQuery3 += " AND AKD_DATA BETWEEN '"+_cFirstD+"' AND '"+_cUltimoD+"' "+ c_ent
_cQuery3 += " AND AKD.D_E_L_E_T_ = ' ' "+ c_ent
_cQuery3 += " AND AK5.D_E_L_E_T_ = ' ' "+ c_ent
_cQuery3 := ChangeQuery(_cQuery3) 

If Select(_cTrab3) > 0
	dbSelectArea(_cTrab3)
	(_cTrab3)->(dbCloseArea())
EndIf

dbUseArea(.T.,"TOPCONN",TcGenQry(,,_cQuery3),_cTrab3,.T.,.T.)       

dbSelectArea(_cTrab3)

If(!Eof())
  While !(_cTrab3)->(Eof())  
		//Aadd(_cArray3,{AllTrim((_cTrab3)->AKD_CO),AllTrim((_cTrab3)->AKD_CC)})
		Aadd(_cArray3,{AllTrim((_cTrab3)->AKD_CO)})
  	DbSkip()
  EndDo 
EndIf  

(_cTrab3)->(dbCloseArea())*/

RestArea(_cArea3)

Return(_cArray3)            


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³OIFR002   ºAutor  ³Leandro Ribeiro     º Data ³  19/06/14   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Função para retornar os saldos das contas                  º±±
±±º          ³                                                            º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP                                                        º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function CONTREAL(_cCO)

Local _cArea4    := GetArea()
Local _cQuery4   := ""
Local _cTrab4    := GetNextAlias() 
Local _cRetSal   := ""                    
Local _cTpSaldo  := MV_PAR01 
Local _cItemCont := _cCO
Local _cOperOrc  := Alltrim(MV_PAR07)
Local _cDataDe   := _cFirstD
Local _cDataAte  := _cUltimoD
                               
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ//
//"Saldo Realizado"			MV_PAR01  //
//"Saldo Orçado"			MV_PAR02  //
//"Item Contabil de?" 		MV_PAR03  //
//"Item Contabil Ate?" 		MV_PAR04  //
//"Centro de Custo de?" 	MV_PAR05  //
//"Centro de Custo ate?" 	MV_PAR06  //
//"Operação Orçamentaria" 	MV_PAR07  //
//"Data" 			   		MV_PAR08  //
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ// 

_cQuery4 := " SELECT AKD_CLVLR,"
Do Case
	Case(_cTpSaldo == 1)  
	_cQuery4 += " SUM(CASE WHEN AKD_TPSALD IN ('PR') THEN AKD_VALOR1 * (CASE WHEN AKD_TIPO='1' THEN 1 ELSE -1 END) ELSE 0 END) as VLTOTAL"+ c_ent
	Case(_cTpSaldo == 2)  
	_cQuery4 += " SUM(CASE WHEN AKD_TPSALD IN ('EM') THEN AKD_VALOR1 * (CASE WHEN AKD_TIPO='1' THEN 1 ELSE -1 END) ELSE 0 END) as VLTOTAL"+ c_ent
	Case(_cTpSaldo == 3)  
	_cQuery4 += " SUM(CASE WHEN AKD_TPSALD IN ('RE') THEN AKD_VALOR1 * (CASE WHEN AKD_TIPO='1' THEN 1 ELSE -1 END) ELSE 0 END) as VLTOTAL"+ c_ent
	Case(_cTpSaldo == 4)  
	_cQuery4 += " SUM(CASE WHEN AKD_TPSALD IN ('PG') THEN AKD_VALOR1 * (CASE WHEN AKD_TIPO='1' THEN 1 ELSE -1 END) ELSE 0 END) as VLTOTAL"+ c_ent
	Case(_cTpSaldo == 5)  
	_cQuery4 += " SUM(CASE WHEN AKD_TPSALD IN ('PR','EM','RE') THEN AKD_VALOR1 * (CASE WHEN AKD_TIPO='1' THEN 1 ELSE -1 END) ELSE 0 END) as VLTOTAL"
	EndCase
_cQuery4 += " FROM "+RETSQLNAME("AKD")+" AKD "+ c_ent
_cQuery4 += " WHERE AKD.AKD_FILIAL = '"+xFilial("AKD")+"'"+ c_ent 
_cQuery4 += " AND AKD_CO = '"+_cItemCont+"'"+ c_ent
_cQuery4 += " AND AKD_DATA  BETWEEN '"+_cDataDe+"' AND '"+_cDataAte+"'"+ c_ent
//_cQuery4 += " AND AKD_OPER = '"+_cOperOrc+"'"+ c_ent  
_cQuery4 += " AND AKD_OPER BETWEEN '"+AllTrim(MV_PAR07)+"' AND '"+AllTrim(MV_PAR08)+"' "+ c_ent
_cQuery4 += " AND AKD_CO BETWEEN '"+MV_PAR10+"' AND '"+MV_PAR11+"' "+ c_ent  
_cQuery4 += " AND AKD_CC BETWEEN '"+MV_PAR05+"' AND '"+MV_PAR06+"' "+ c_ent
_cQuery4 += " AND AKD_STATUS = '1'"+ c_ent
_cQuery4 += " AND AKD.D_E_L_E_T_ = ' '"+ c_ent
_cQuery4 += " GROUP BY AKD_CLVLR"+ c_ent
_cQuery4 += " ORDER BY AKD_CLVLR"+ c_ent
_cQuery4 := ChangeQuery(_cQuery4)

If Select(_cTrab4) > 0
	dbSelectArea(_cTrab4)
	(_cTrab4)->(dbCloseArea())
EndIf

dbUseArea(.T.,"TOPCONN",TcGenQry(,,_cQuery4),_cTrab4,.T.,.T.)       

dbSelectArea(_cTrab4)

If(!Eof())
  While !(_cTrab4)->(Eof())  
		_cRetSal := STR((_cTrab4)->VLTOTAL)
  	DbSkip()
  EndDo 
Else
   _cRetSal := "0.00"
EndIf                   

(_cTrab4)->(dbCloseArea()) 

RestArea(_cArea4)

Return(_cRetSal)      

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³OIFR002   ºAutor  ³Leandro Ribeiro     º Data ³  19/06/14   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Função para retornar os saldos das contas                  º±±
±±º          ³                                                            º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP                                                        º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function CONTORC(_cCO)

Local _cArea6    := GetArea()
Local _cQuery6   := ""
Local _cTrab6    := GetNextAlias() 
Local _cRetSal   := ""                    
Local _cTpSaldo  := MV_PAR02 
Local _cItemCont := _cCO
Local _cOperOrc  := Alltrim(MV_PAR07)
Local _cDataDe   := _cFirstD
Local _cDataAte  := _cUltimoD
                               
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ//
//"Saldo Realizado"			MV_PAR01  //
//"Saldo Orçado"			MV_PAR02  //
//"Item Contabil de?" 		MV_PAR03  //
//"Item Contabil Ate?" 		MV_PAR04  //
//"Centro de Custo de?" 	MV_PAR05  //
//"Centro de Custo ate?" 	MV_PAR06  //
//"Operação Orçamentaria" 	MV_PAR07  //
//"Data" 			   		MV_PAR08  //
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ// 

_cQuery6 := " SELECT AKD_CLVLR,"
Do Case
	Case(_cTpSaldo == 1)  
	_cQuery6 += " SUM(CASE WHEN AKD_TPSALD IN ('CT') THEN AKD_VALOR1 * (CASE WHEN AKD_TIPO='1' THEN 1 ELSE -1 END) ELSE 0 END) as VLTOTAL"+ c_ent
	Case(_cTpSaldo == 2)  
	_cQuery6 += " SUM(CASE WHEN AKD_TPSALD IN ('OR') THEN AKD_VALOR1 * (CASE WHEN AKD_TIPO='1' THEN 1 ELSE -1 END) ELSE 0 END) as VLTOTAL"+ c_ent
	Case(_cTpSaldo == 3)  
	_cQuery6 += " SUM(CASE WHEN AKD_TPSALD IN ('OR','CT') THEN AKD_VALOR1 * (CASE WHEN AKD_TIPO='1' THEN 1 ELSE -1 END) ELSE 0 END) as VLTOTAL"
	EndCase
_cQuery6 += " FROM "+RETSQLNAME("AKD")+" AKD "+ c_ent
_cQuery6 += " WHERE AKD.AKD_FILIAL = '"+xFilial("AKD")+"'"+ c_ent 
_cQuery6 += " AND AKD_CO = '"+_cItemCont+"'"+ c_ent
_cQuery6 += " AND AKD_DATA  BETWEEN '"+_cDataDe+"' AND '"+_cDataAte+"'"+ c_ent
//_cQuery6 += " AND AKD_OPER = '"+_cOperOrc+"'"+ c_ent  
_cQuery6 += " AND AKD_OPER BETWEEN '"+AllTrim(MV_PAR07)+"' AND '"+AllTrim(MV_PAR08)+"' "+ c_ent
_cQuery6 += " AND AKD_CO BETWEEN '"+MV_PAR10+"' AND '"+MV_PAR11+"' "+ c_ent  
_cQuery6 += " AND AKD_CC BETWEEN '"+MV_PAR05+"' AND '"+MV_PAR06+"' "+ c_ent
_cQuery6 += " AND AKD_STATUS = '1'"+ c_ent
_cQuery6 += " AND AKD.D_E_L_E_T_ = ' '"+ c_ent
_cQuery6 += " GROUP BY AKD_CLVLR"+ c_ent
_cQuery6 += " ORDER BY AKD_CLVLR"+ c_ent
_cQuery6 := ChangeQuery(_cQuery6)

If Select(_cTrab6) > 0
	dbSelectArea(_cTrab6)
	(_cTrab6)->(dbCloseArea())
EndIf

dbUseArea(.T.,"TOPCONN",TcGenQry(,,_cQuery6),_cTrab6,.T.,.T.)       

dbSelectArea(_cTrab6)

If(!Eof())
  While !(_cTrab6)->(Eof())  
		_cRetSal := STR((_cTrab6)->VLTOTAL)
  	DbSkip()
  EndDo 
Else
   _cRetSal := "0.00"
EndIf                   

(_cTrab6)->(dbCloseArea()) 

RestArea(_cArea6)

Return(_cRetSal)    

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³OIFR002   ºAutor  ³Leandro Ribeiro     º Data ³  19/06/14   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Função para retornar os saldos das contas                  º±±
±±º          ³                                                            º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP                                                        º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function CONTSALD(_cCO)

Local _cArea7    := GetArea()
Local _cQuery7   := ""
Local _cTrab7    := GetNextAlias() 
Local _cRetSal   := ""                    
Local _cTpSaldo1 := MV_PAR01 
Local _cTpSaldo2 := MV_PAR02 
Local _cItemCont := _cCO
Local _cOperOrc  := Alltrim(MV_PAR07)
Local _cDataDe   := _cFirstD
Local _cDataAte  := _cUltimoD 
Local _cTSaldo1  := ""
Local _cTSaldo2  := ""
                               
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ//
//"Saldo Realizado"			MV_PAR01  //
//"Saldo Orçado"			MV_PAR02  //
//"Item Contabil de?" 		MV_PAR03  //
//"Item Contabil Ate?" 		MV_PAR04  //
//"Centro de Custo de?" 	MV_PAR05  //
//"Centro de Custo ate?" 	MV_PAR06  //
//"Operação Orçamentaria" 	MV_PAR07  //
//"Data" 			   		MV_PAR08  //
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ// 

Do Case
	Case(_cTpSaldo1 == 1)  
		_cTSaldo1 := "('PR')"
	Case(_cTpSaldo1 == 2)  
		_cTSaldo1 := "('EM')"
	Case(_cTpSaldo1 == 3)  
		_cTSaldo1 := "('RE')"
	Case(_cTpSaldo1 == 4)  
		_cTSaldo1 := "('PG')"
	Case(_cTpSaldo1 == 5)  
		_cTSaldo1 := "('PR','EM','RE')"
EndCase   
Do Case
	Case(_cTpSaldo2 == 1)  
		_cTSaldo2 := "('CT')"
	Case(_cTpSaldo2 == 2)  
		_cTSaldo2 := "('OR')"
	Case(_cTpSaldo2 == 3)  
		_cTSaldo2 := "('OR','CT')"
EndCase


_cQuery7 := " SELECT AKD_CLVLR,"
//_cQuery7 += " SUM(CASE WHEN AKD_TPSALD IN ('OR','CT') THEN AKD_VALOR1 * (CASE WHEN AKD_TIPO='1' THEN 1 ELSE -1 END) ELSE 0 END) - SUM(CASE WHEN AKD_TPSALD IN ('PR','EM','RE','PG') THEN AKD_VALOR1 * (CASE WHEN AKD_TIPO='1' THEN 1 ELSE -1 END) ELSE 0 END) as SALDO
_cQuery7 += " SUM(CASE WHEN AKD_TPSALD IN " +_cTSaldo2+ " THEN AKD_VALOR1 * (CASE WHEN AKD_TIPO='1' THEN 1 ELSE -1 END) ELSE 0 END) - SUM(CASE WHEN AKD_TPSALD IN " +_cTSaldo1+ " THEN AKD_VALOR1 * (CASE WHEN AKD_TIPO='1' THEN 1 ELSE -1 END) ELSE 0 END) as SALDO
_cQuery7 += " FROM "+RETSQLNAME("AKD")+" AKD "+ c_ent
_cQuery7 += " WHERE AKD.AKD_FILIAL = '"+xFilial("AKD")+"'"+ c_ent 
_cQuery7 += " AND AKD_CO = '"+_cItemCont+"'"+ c_ent
_cQuery7 += " AND AKD_DATA  BETWEEN '"+_cDataDe+"' AND '"+_cDataAte+"'"+ c_ent
//_cQuery7 += " AND AKD_OPER = '"+_cOperOrc+"'"+ c_ent  
_cQuery7 += " AND AKD_OPER BETWEEN '"+AllTrim(MV_PAR07)+"' AND '"+AllTrim(MV_PAR08)+"' "+ c_ent
_cQuery7 += " AND AKD_CO BETWEEN '"+MV_PAR10+"' AND '"+MV_PAR11+"' "+ c_ent  
_cQuery7 += " AND AKD_CC BETWEEN '"+MV_PAR05+"' AND '"+MV_PAR06+"' "+ c_ent
_cQuery7 += " AND AKD_STATUS = '1'"+ c_ent
_cQuery7 += " AND AKD.D_E_L_E_T_ = ' '"+ c_ent
_cQuery7 += " GROUP BY AKD_CLVLR"+ c_ent
_cQuery7 += " ORDER BY AKD_CLVLR"+ c_ent
_cQuery7 := ChangeQuery(_cQuery7)

If Select(_cTrab7) > 0
	dbSelectArea(_cTrab7)
	(_cTrab7)->(dbCloseArea())
EndIf

dbUseArea(.T.,"TOPCONN",TcGenQry(,,_cQuery7),_cTrab7,.T.,.T.)       

dbSelectArea(_cTrab7)

If(!Eof())
  While !(_cTrab7)->(Eof())  
		_cRetSal := STR((_cTrab7)->SALDO * -1)
  	DbSkip()
  EndDo 
Else
   _cRetSal := "0.00"
EndIf                   

(_cTrab7)->(dbCloseArea()) 

RestArea(_cArea7)

Return(_cRetSal) 

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³OIFR002   ºAutor  ³Leandro Ribeiro     º Data ³  19/06/14   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Função para verificar a existencia de saldo na contas      º±±
±±º          ³                                                            º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP                                                        º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function VERISALDO(_cCO,_cCC)

Local _cArea5    := GetArea()
Local _cQuery5   := ""
Local _cTrab5    := GetNextAlias() 
Local _cRetSal   := ""                    
Local _cTpSaldo  := AllTrim(MV_PAR01) 
Local _cItemCont := _cCO
Local _cCCusto   := _cCC
Local _cOperOrc  := Alltrim(MV_PAR07)
Local _cDataDe   := _cFirstD           
Local _cDataAte  := _cUltimoD             

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ//
//"Saldo Realizado"			MV_PAR01  //
//"Saldo Orçado"			MV_PAR02  //
//"Item Contabil de?" 		MV_PAR03  //
//"Item Contabil Ate?" 		MV_PAR04  //
//"Centro de Custo de?" 	MV_PAR05  //
//"Centro de Custo ate?" 	MV_PAR06  //
//"Operação Orçamentaria" 	MV_PAR07  //
//"Data" 			   		MV_PAR08  //
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ//  


_cQuery5 := " SELECT AKD_CLVLR,"
Do Case
	Case(_cTpSaldo == 'PR')  
	_cQuery5 += " SUM(CASE WHEN AKD_TPSALD IN ('PR') THEN AKD_VALOR1 * (CASE WHEN AKD_TIPO='1' THEN 1 ELSE -1 END) ELSE 0 END) as VLTOTAL"+ c_ent
	Case(_cTpSaldo == 'EM')  
	_cQuery5 += " SUM(CASE WHEN AKD_TPSALD IN ('EM') THEN AKD_VALOR1 * (CASE WHEN AKD_TIPO='1' THEN 1 ELSE -1 END) ELSE 0 END) as VLTOTAL"+ c_ent
	Case(_cTpSaldo == 'RE')  
	_cQuery5 += " SUM(CASE WHEN AKD_TPSALD IN ('RE') THEN AKD_VALOR1 * (CASE WHEN AKD_TIPO='1' THEN 1 ELSE -1 END) ELSE 0 END) as VLTOTAL"+ c_ent
	Case(_cTpSaldo == 'OR')  
	_cQuery5 += " SUM(CASE WHEN AKD_TPSALD IN ('OR') THEN AKD_VALOR1 * (CASE WHEN AKD_TIPO='1' THEN 1 ELSE -1 END) ELSE 0 END) as VLTOTAL"+ c_ent
	Case(_cTpSaldo == 'CT')  
	_cQuery5 += " SUM(CASE WHEN AKD_TPSALD IN ('CT') THEN AKD_VALOR1 * (CASE WHEN AKD_TIPO='1' THEN 1 ELSE -1 END) ELSE 0 END) as VLTOTAL"+ c_ent
	Case(_cTpSaldo == 'PG')  
	_cQuery5 += " SUM(CASE WHEN AKD_TPSALD IN ('PG') THEN AKD_VALOR1 * (CASE WHEN AKD_TIPO='1' THEN 1 ELSE -1 END) ELSE 0 END) as VLTOTAL"+ c_ent
	Case(_cTpSaldo == 'PR+EM+RE')  
	_cQuery5 += " SUM(CASE WHEN AKD_TPSALD IN ('PR','EM','RE') THEN AKD_VALOR1 * (CASE WHEN AKD_TIPO='1' THEN 1 ELSE -1 END) ELSE 0 END) as VLTOTAL"
	EndCase
_cQuery5 += " FROM "+RETSQLNAME("AKD")+" AKD "+ c_ent
_cQuery5 += " WHERE AKD.AKD_FILIAL = '"+xFilial("AKD")+"'"+ c_ent /*"+xFilial("AKD")+"*/
//_cQuery4 += " AND AKD_ITCTB = '"+_cItemCont+"'"+ c_ent
_cQuery5 += " AND AKD_CO = '"+_cItemCont+"'"+ c_ent
_cQuery5 += " AND AKD_DATA  BETWEEN '"+_cDataDe+"' AND '"+_cDataAte+"'"+ c_ent
//_cQuery5 += " AND AKD_OPER = '"+_cOperOrc+"'"+ c_ent  
_cQuery5 += " AND AKD_OPER BETWEEN '"+AllTrim(MV_PAR07)+"' AND '"+AllTrim(MV_PAR08)+"' "+ c_ent
_cQuery5 += " AND AKD_CO BETWEEN '"+MV_PAR10+"' AND '"+MV_PAR11+"' "+ c_ent  
_cQuery5 += " AND AKD_CC = '"+_cCCusto+"'"+ c_ent
_cQuery5 += " AND AKD_STATUS = '1'"+ c_ent
_cQuery5 += " AND AKD.D_E_L_E_T_ = ' '"+ c_ent
_cQuery5 += " GROUP BY AKD_CLVLR"+ c_ent
_cQuery5 += " ORDER BY AKD_CLVLR"+ c_ent
_cQuery5 := ChangeQuery(_cQuery5)

If Select(_cTrab5) > 0
	dbSelectArea(_cTrab5)
	(_cTrab5)->(dbCloseArea())
EndIf

dbUseArea(.T.,"TOPCONN",TcGenQry(,,_cQuery5),_cTrab5,.T.,.T.)       

dbSelectArea(_cTrab5)

If(!Eof())
  While !(_cTrab5)->(Eof())  
		//_cRetSal := Alltrim(Transform((_cTrab4)->VLTOTAL, "@E 999999999.99" ))
		_cRetSal := STR((_cTrab5)->VLTOTAL)
  	DbSkip()
  EndDo 
Else
   _cRetSal := "0"
EndIf                   

(_cTrab5)->(dbCloseArea()) 

RestArea(_cArea5)

Return(_cRetSal)


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³OIFR002   ºAutor  ³Leandro Ribeiro     º Data ³  20/06/14   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Função para calcular a posição das contas superiores para  º±±
±±º          ³ adicionar no rodapé da planilha.                           º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP                                                        º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function LINHACALC(_cLinhas)

Local _cRetLin := ""
Local _cLinFin := 0  
Local _cLinSom := {}
Local _cNegar  := {}
Local _cInvert := {}

For _aa := Len(_cLinhas) to 1 Step -1
 	Aadd(_cInvert, _cLinhas[_aa])
Next _aa
                  
_cLinhas := _cInvert        
         
For _ee := 1 to Len(_cLinhas)
	Aadd(_cLinSom, _cLinFin := (_cLinFin + _cLinhas[_ee]))
Next _ee                                                     

For _ii := 1 to Len(_cLinSom)
	Aadd(_cNegar,_cLinSom[_ii] * -1)
Next _ii
     
_cLinSom := _cNegar
                 
If(Len(_cLinSom) == 1)          
	For _ss := 1 to Len(_cLinSom)
		_cRetLin := "SUM(R["+Alltrim(STR(_cLinSom[_ss]))+"]C)"
	Next _ss
Else
	_cRetLin := "SUM("   
	For _ss := 1 to Len(_cLinSom)
		If!(_ss == Len(_cLinSom))
			_cRetLin += "R["+Alltrim(STR(_cLinSom[_ss]))+"]C,"
		Else
			_cRetLin += "R["+Alltrim(STR(_cLinSom[_ss]))+"]C"
		EndIf
	Next _ss 
	_cRetLin += ")"   
EndIf

Return(_cRetLin)


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³OIFR002   ºAutor  ³Microsiga           º Data ³  06/21/14   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³                                                            º±±
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
±±ºPrograma  ³OIFR002   ºAutor  ³Leandro Ribeiro     º Data ³  28/05/14   º±±
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
   
Return(_sRet)


