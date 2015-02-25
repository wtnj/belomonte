#INCLUDE "Protheus.ch"
#INCLUDE "TopConn.ch"
#INCLUDE "rwmake.ch" 

#DEFINE c_ent CHR(13)+CHR(10)

#DEFINE _OPC_cGETFILE ( GETF_RETDIRECTORY + GETF_ONLYSERVER )


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³OIFR001   ºAutor  ³Leandro Ribeiro     º Data ³  13/06/14   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Programa para gerar planilha de conta com movimentações    º±±
±±º          ³ na tabela AKD de acordo com a parametrização do usuario.   º±± 
±±º          ³ Planilha Conta Orçamentaria Mes a Mes.					  º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP                                                        º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

User Function OIFR001()  

Local _aArea := GetArea() 
Local cPerg  := Padr("OIFR001",Len(SX1->X1_GRUPO))      	             
                                                                                      //AL2A                                                                                                     ENG
//PutSx1(cPerg,"01","Tipo de Saldo"         , "", "", "mv_ch1" , "C", 08,0 ,0, "C", "", ""      ,"", "", "MV_PAR01", "PG" , "" , "" , "", "RE"      ,"","","EM","","","CT","","","","","",""/*,"","",""*///)     
PutSx1(cPerg,"01","Tipo de Saldo"           , "", "", "mv_ch1" , "C", TAMSX3("AKD_TPSALD")[1],0 ,0, "G", "", "AL2A"    ,"", "", "MV_PAR01", "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"")
PutSx1(cPerg,"02","Item Contabil de?"       , "", "", "mv_ch2" , "C", TAMSX3("AKD_ITCTB")[1] ,0 ,0, "G", "", "CTD"     ,"", "", "MV_PAR02", "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"")
PutSx1(cPerg,"03","Item Contabil Ate?"      , "", "", "mv_ch3" , "C", TAMSX3("AKD_ITCTB")[1] ,0 ,0, "G", "", "CTD"     ,"", "", "MV_PAR03", "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"")
PutSx1(cPerg,"04","Centro de Custo de?"     , "", "", "mv_ch4" , "C", TAMSX3("AKD_CC")[1]    ,0, 0, "G", "", "CTT"     ,"", "", "MV_PAR04", "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"") 
PutSx1(cPerg,"05","Centro de Custo ate?"    , "", "", "mv_ch5" , "C", TAMSX3("AKD_CC")[1]    ,0, 0, "G", "", "CTT"     ,"", "", "MV_PAR05", "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"") 
//PutSx1(cPerg,"06","Operação Orçamentaria" , "", "", "mv_ch6" , "C", 1                      ,0, 0, "C", "", " "       ,"", "", "MV_PAR06", "1" , "" , "" , "" , "2" , "" , "" , "3" , "" , "" , "" , "","","","","","","","","","",""   ,"")
PutSx1(cPerg,"06","Oper. Orçamentaria de?"  , "", "", "mv_ch6" , "C", TAMSX3("AKF_CODIGO")[1]   ,0, 0, "G", "", "AKF" ,"", "", "MV_PAR06" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"")
PutSx1(cPerg,"07","Oper. Orçamentaria ate?" , "", "", "mv_ch7" , "C", TAMSX3("AKF_CODIGO")[1]   ,0, 0, "G", "", "AKF" ,"", "", "MV_PAR07" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"")
PutSx1(cPerg,"08","Data de"                 , "", "", "mv_ch8" , "D", TAMSX3("AKD_DATA")[1]  ,0, 0, "G", "", ""        ,"", "", "MV_PAR08", "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"")
PutSx1(cPerg,"09","Data Ate"                , "", "", "mv_ch9" , "D", TAMSX3("AKD_DATA")[1]  ,0, 0, "G", "", ""        ,"", "", "MV_PAR09", "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"")
PutSx1(cPerg,"10","Conta Orçamentaria de?"    , "", "", "mv_ch10" , "C", TAMSX3("AK5_CODIGO")[1]   ,0, 0, "G", "", "AK5"     ,"", "", "MV_PAR10" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"") 
PutSx1(cPerg,"11","Conta Orçamentaria ate?"   , "", "", "mv_ch11" , "C", TAMSX3("AK5_CODIGO")[1]   ,0, 0, "G", "", "AK5"     ,"", "", "MV_PAR11" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "" , "","","","","","","","","","",""   ,"") 
       

If Pergunte(cPerg,.T.)
	
	Processa({|| GOIFR001() },"Aguarde...")
	
EndIf

RestArea(_aArea)
                                               
Return()


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³GOIFR001   ºAutor  ³Leandro Ribeiro    º Data ³  13/06/14   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Função para gerar a planilha de excel                      º±±
±±º          ³                                                            º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP                                                        º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

Static Function GOIFR001()

Local _cOIFR001 := GetArea()
//Local _cFile    := "OIFR001" + DToS( Date() ) + ".xml"                      
Local _cFile    := "OIFR001.xml"                      
Local _cPath    := GetTempPath() //Pega a Pasta de Arquivo temporário                 

Local _cQuery1  := "" 
Local _cTrab1  	:= GetNextAlias() 
Local _cTrab2  	:= GetNextAlias()
Local _cArray1  := {} //Retorna as Contas Orçamentarias Superior 
Local _cArray2	:= {} //Retorna as Contas Orçamentarias  
Local _cArray5  := {}
Local _cArray6  := {} 
Local _cArray7  := {}
Local _cArray8  := {}
Local _cArray9  := {}
Local _cVerSald := "" 
Local _cFlag1   := .F. 
Local _cCont    := 0
Local _cRetSoma := ""

Private cPath    := "" 
Private _nHandle 
                      
dbSelectArea("AK5")
dbSetOrder(1)      
dbSelectArea("AKD")
dbSetOrder(1)   

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
//"Tipo de Saldo" 			MV_PAR01  //
//"Item Contabil de?" 		MV_PAR02  //
//"Item Contabil Ate?" 		MV_PAR03  //
//"Centro de Custo de?" 	MV_PAR04  //
//"Centro de Custo ate?" 	MV_PAR05  //
//"Oper. Orçamentaria de?" 	MV_PAR06  //
//"Oper. Orçamentaria Ate?" MV_PAR07  //
//"Data de" 				MV_PAR08  //
//"Data Ate" 				MV_PAR09  //
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ//

_cQuery1 := " SELECT DISTINCT AK5_COSUP " + c_ent
_cQuery1 += " FROM "+RETSQLNAME("AKD")+" AKD "+ c_ent
_cQuery1 += " INNER JOIN "+RETSQLNAME("AK5")+" AK5 "+ c_ent
_cQuery1 += " ON AKD_FILIAL = AK5_FILIAL "+ c_ent 
_cQuery1 += " AND AKD_CO = AK5_CODIGO "+ c_ent
_cQuery1 += " WHERE "+ c_ent
_cQuery1 += " AKD_TPSALD = '"+AllTrim(MV_PAR01)+"' "+ c_ent
_cQuery1 += " AND AKD_CC BETWEEN '"+AllTrim(MV_PAR04)+"' AND '"+AllTrim(MV_PAR05)+"' "+ c_ent
_cQuery1 += " AND AKD_ITCTB BETWEEN '"+AllTrim(MV_PAR02)+"' AND '"+AllTrim(MV_PAR03)+"' "+ c_ent
//_cQuery1 += " AND AKD_OPER = '"+Alltrim(MV_PAR06)+"' "+ c_ent                                 
_cQuery1 += " AND AKD_OPER BETWEEN '"+AllTrim(MV_PAR06)+"' AND '"+AllTrim(MV_PAR07)+"' "+ c_ent
_cQuery1 += " AND AKD_DATA BETWEEN '"+DTOS(MV_PAR08)+"' AND '"+DTOS(MV_PAR09)+"' "+ c_ent      
_cQuery1 += " AND AKD_CO BETWEEN '"+MV_PAR10+"' AND '"+MV_PAR11+"' "+ c_ent  
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

If(EMPTY(_cArray1))
   Aviso("Aviso","Nenhum dado encontrado com os parametros passados!",{"OK"}) 
   		U_OIFR001()
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
fWrite(_nHandle,'  <LastSaved>2014-06-19T21:55:35Z</LastSaved>'+ c_ent)
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
fWrite(_nHandle,'   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>'+ c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Color="#FFFFFF"/>'+ c_ent)
fWrite(_nHandle,'   <Interior ss:Color="#7F7F7F" ss:Pattern="Solid"/>'+ c_ent)
fWrite(_nHandle,'  </Style>'+ c_ent)
fWrite(_nHandle,'  <Style ss:ID="s70">'+ c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>'+ c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Color="#000000"/>'+ c_ent)
fWrite(_nHandle,'   <Interior/>'+ c_ent)
fWrite(_nHandle,'  </Style>'+ c_ent)
fWrite(_nHandle,'  <Style ss:ID="s71">'+ c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>'+ c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Color="#000000"/>'+ c_ent)
fWrite(_nHandle,'   <Interior/>'+ c_ent)
fWrite(_nHandle,'  </Style>'+ c_ent)
fWrite(_nHandle,'  <Style ss:ID="s72">'+ c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>'+ c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Color="#000000" ss:Bold="1"/>'+ c_ent)
fWrite(_nHandle,'   <Interior ss:Color="#D8D8D8" ss:Pattern="Solid"/>'+ c_ent)
fWrite(_nHandle,'  </Style>'+ c_ent) 
fWrite(_nHandle,'    <Style ss:ID="s73">'+ c_ent) 
fWrite(_nHandle,'     <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>'+ c_ent) 
fWrite(_nHandle,'     <Font ss:FontName="Calibri" x:Family="Swiss" ss:Color="#993300"/>'+ c_ent) 
fWrite(_nHandle,'     <Interior ss:Color="#D8D8D8" ss:Pattern="Solid"/>'+ c_ent) 
fWrite(_nHandle,'    </Style>'+ c_ent) 
fWrite(_nHandle,'  <Style ss:ID="s74">'+ c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom" ss:Indent="2"/>'+ c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Color="#000000"/>'+ c_ent)
fWrite(_nHandle,'   <Interior/>'+ c_ent)
fWrite(_nHandle,'  </Style>'+ c_ent)
fWrite(_nHandle,'  <Style ss:ID="s75">'+ c_ent)
fWrite(_nHandle,'     <Alignment ss:Horizontal="Left" ss:Vertical="Bottom" ss:Indent="2"/>'+ c_ent)
fWrite(_nHandle,'     <Font ss:FontName="Calibri" x:Family="Swiss" ss:Color="#000000"/>'+ c_ent)
fWrite(_nHandle,'     <Interior/>'+ c_ent)
fWrite(_nHandle,'    </Style>'+ c_ent)
fWrite(_nHandle,'  <Style ss:ID="s76">'+ c_ent)
fWrite(_nHandle,'     <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>'+ c_ent)
fWrite(_nHandle,'     <Font ss:FontName="Calibri" x:Family="Swiss" ss:Color="#000000"/>'+ c_ent)
fWrite(_nHandle,'     <Interior/>'+ c_ent)
fWrite(_nHandle,'     <NumberFormat ss:Format="Fixed"/>'+ c_ent)
fWrite(_nHandle,'    </Style>'+ c_ent)
fWrite(_nHandle,'   <Style ss:ID="s77">'+ c_ent)
fWrite(_nHandle,'     <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>'+ c_ent)
fWrite(_nHandle,'     <Font ss:FontName="Calibri" x:Family="Swiss" ss:Color="#FFFFFF" ss:Bold="1"/>'+ c_ent)
fWrite(_nHandle,'     <Interior ss:Color="#7F7F7F" ss:Pattern="Solid"/>'+ c_ent)
fWrite(_nHandle,'    </Style>'+ c_ent)
fWrite(_nHandle,'  <Style ss:ID="s78">'+ c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>'+ c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Color="#993300"/>'+ c_ent)
fWrite(_nHandle,'   <Interior ss:Color="#D8D8D8" ss:Pattern="Solid"/>'+ c_ent)
fWrite(_nHandle,'   <NumberFormat ss:Format="Fixed"/>'+ c_ent)
fWrite(_nHandle,'  </Style>'+ c_ent)
fWrite(_nHandle,'  <Style ss:ID="s79">'+ c_ent)
fWrite(_nHandle,'   <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>'+ c_ent)
fWrite(_nHandle,'   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Color="#FFFFFF"/>'+ c_ent)
fWrite(_nHandle,'   <Interior ss:Color="#7F7F7F" ss:Pattern="Solid"/>'+ c_ent)
fWrite(_nHandle,'   <NumberFormat ss:Format="Fixed"/>'+ c_ent)
fWrite(_nHandle,'  </Style>'+ c_ent)
fWrite(_nHandle,' </Styles>'+ c_ent) 
fWrite(_nHandle,' <Worksheet ss:Name="conta mes a mes">'+ c_ent)
fWrite(_nHandle,'  <Table ss:ExpandedColumnCount="16" ss:ExpandedRowCount="10000" x:FullColumns="1"'+ c_ent)
fWrite(_nHandle,'   x:FullRows="1" ss:DefaultRowHeight="15">'+ c_ent)
fWrite(_nHandle,'   <Column ss:AutoFitWidth="0" ss:Width="15.75"/>'+ c_ent)
fWrite(_nHandle,'   <Column ss:AutoFitWidth="0" ss:Width="182.25"/>'+ c_ent)
fWrite(_nHandle,'   <Column ss:AutoFitWidth="0" ss:Width="72"/>'+ c_ent)
fWrite(_nHandle,'   <Column ss:StyleID="s62" ss:AutoFitWidth="0" ss:Width="72"/>'+ c_ent)
fWrite(_nHandle,'   <Column ss:StyleID="s62" ss:AutoFitWidth="0" ss:Width="48.75" ss:Span="11"/>'+ c_ent)
fWrite(_nHandle,'   <Row ss:AutoFitHeight="0">'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s63"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s63"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s63"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
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
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'   </Row>'+ c_ent)
fWrite(_nHandle,'   <Row ss:AutoFitHeight="0">'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s63"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s65"><Data ss:Type="String">Diretoria de Planejamento e Desempenho</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s65"/>'+ c_ent)
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
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'   </Row>'+ c_ent)
fWrite(_nHandle,'   <Row ss:AutoFitHeight="0">'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s63"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s65"><Data ss:Type="String">Estrutura Oi Futuro</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s65"/>'+ c_ent)
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
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'   </Row>'+ c_ent)
fWrite(_nHandle,'   <Row ss:AutoFitHeight="0">'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s63"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s67"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s67"/>'+ c_ent)
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
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s64"/>'+ c_ent) 
fWrite(_nHandle,'   </Row>'+ c_ent)
fWrite(_nHandle,'   <Row ss:AutoFitHeight="0">'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s63"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s68"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s68"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s69"><Data ss:Type="String">Jan</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s69"><Data ss:Type="String">Fev</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s69"><Data ss:Type="String">Mar</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s69"><Data ss:Type="String">Abr</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s69"><Data ss:Type="String">Mai</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s69"><Data ss:Type="String">Jun</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s69"><Data ss:Type="String">Jul</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s69"><Data ss:Type="String">Ago</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s69"><Data ss:Type="String">Set</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s69"><Data ss:Type="String">Out</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s69"><Data ss:Type="String">Nov</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s69"><Data ss:Type="String">Dez</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s69"><Data ss:Type="String">Total</Data></Cell>'+ c_ent)
fWrite(_nHandle,'   </Row>'+ c_ent)
fWrite(_nHandle,'   <Row ss:AutoFitHeight="0" ss:Height="4.5">'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s63"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s70"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s70"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
fWrite(_nHandle,'   </Row>'+ c_ent)  
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

For _xx := 1 to Len(_cArray1)    
	
	_cArray6 := RETCONTA(_cArray1[_xx]) 
	
	_cFlag1 := .F.    
	
		For _rr := 1 to Len(_cArray6)
	
			_cVerSald := VERISALDO(_cArray6[_rr][1]) 
	
				If!(AllTrim(_cVerSald) == "0" .OR. EMPTY(AllTrim(_cVerSald)))
	
					Aadd(_cArray7,_cArray6[_rr][1])	
	
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
_cArray2 := _cArray7     

If(EMPTY(_cArray1))
   Aviso("Aviso","Nenhum dado encontrado com os parametros passados!",{"OK"}) 
   		U_OIFR001()
   Return()
EndIf

For _cc := 1 to Len(_cArray1)         

		_cArray2 := RETCONTA(_cArray1[_cc])

		fWrite(_nHandle,'   <Row ss:AutoFitHeight="0">'+ c_ent)						 
		fWrite(_nHandle,'    <Cell ss:StyleID="s63"/>'+ c_ent)
		//fWrite(_nHandle,'    <Cell ss:StyleID="s72"><Data ss:Type="String">'+Alltrim(GetAdvFVal("AK5","AK5_DESCRI",xFilial("AK5")+PADR(_cArray1[_cc],TAMSX3("AK5_CODIGO")[1]),1))+'</Data></Cell>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s72"><Data ss:Type="String">'+TiraGraf(Alltrim(POSICIONE("AK5",1,xFilial("AK5")+PADR(_cArray1[_cc],TAMSX3("AK5_CODIGO")[1]),"AK5_DESCRI")))+'</Data></Cell>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s72"><Data ss:Type="String">'+Alltrim(_cArray1[_cc])+'</Data></Cell>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s78" ss:Formula="=SUM(R[1]C:R['+Alltrim(STR(Len(_cArray2)))+']C)"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s78" ss:Formula="=SUM(R[1]C:R['+Alltrim(STR(Len(_cArray2)))+']C)"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s78" ss:Formula="=SUM(R[1]C:R['+Alltrim(STR(Len(_cArray2)))+']C)"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s78" ss:Formula="=SUM(R[1]C:R['+Alltrim(STR(Len(_cArray2)))+']C)"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s78" ss:Formula="=SUM(R[1]C:R['+Alltrim(STR(Len(_cArray2)))+']C)"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s78" ss:Formula="=SUM(R[1]C:R['+Alltrim(STR(Len(_cArray2)))+']C)"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s78" ss:Formula="=SUM(R[1]C:R['+Alltrim(STR(Len(_cArray2)))+']C)"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s78" ss:Formula="=SUM(R[1]C:R['+Alltrim(STR(Len(_cArray2)))+']C)"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s78" ss:Formula="=SUM(R[1]C:R['+Alltrim(STR(Len(_cArray2)))+']C)"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s78" ss:Formula="=SUM(R[1]C:R['+Alltrim(STR(Len(_cArray2)))+']C)"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s78" ss:Formula="=SUM(R[1]C:R['+Alltrim(STR(Len(_cArray2)))+']C)"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s78" ss:Formula="=SUM(R[1]C:R['+Alltrim(STR(Len(_cArray2)))+']C)"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s78" ss:Formula="=SUM(RC[-12]:RC[-1])"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
		fWrite(_nHandle,'   </Row>'+ c_ent)                                                                                 
		
		_cCont := _cCont + 2
		
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// CONTAS ORÇAMENTARIAS
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////         
		
			For _zz := 1 to Len(_cArray2)
										
				fWrite(_nHandle,'   <Row ss:AutoFitHeight="0">'+ c_ent)
				fWrite(_nHandle,'    <Cell ss:StyleID="s63"/>'+ c_ent)
				//fWrite(_nHandle,'    <Cell ss:StyleID="s74"><Data ss:Type="String">'+Alltrim(GetAdvfval("AK5","AK5_DESCRI",xFilial("AK5")+PADR(_cArray2[_zz],TAMSX3("AK5_CODIGO")[1]),1))+'</Data></Cell>'+ c_ent)
				fWrite(_nHandle,'    <Cell ss:StyleID="s75"><Data ss:Type="String">'+TiraGraf(Alltrim(POSICIONE("AK5",1,xFilial("AK5")+PADR(_cArray2[_zz][1],TAMSX3("AK5_CODIGO")[1]),"AK5_DESCRI")))+'</Data></Cell>'+ c_ent)
				fWrite(_nHandle,'    <Cell ss:StyleID="s75"><Data ss:Type="String">'+Alltrim(_cArray2[_zz][1])+'</Data></Cell>'+ c_ent)               
				
				Aadd(_cArray5,AllTrim(RETSALDO(_cArray2[_zz][1],"1")))
				Aadd(_cArray5,AllTrim(RETSALDO(_cArray2[_zz][1],"2")))
				Aadd(_cArray5,AllTrim(RETSALDO(_cArray2[_zz][1],"3")))
				Aadd(_cArray5,AllTrim(RETSALDO(_cArray2[_zz][1],"4")))
				Aadd(_cArray5,AllTrim(RETSALDO(_cArray2[_zz][1],"5")))
				Aadd(_cArray5,AllTrim(RETSALDO(_cArray2[_zz][1],"6")))
				Aadd(_cArray5,AllTrim(RETSALDO(_cArray2[_zz][1],"7")))
				Aadd(_cArray5,AllTrim(RETSALDO(_cArray2[_zz][1],"8")))
				Aadd(_cArray5,AllTrim(RETSALDO(_cArray2[_zz][1],"9")))
				Aadd(_cArray5,AllTrim(RETSALDO(_cArray2[_zz][1],"10")))
				Aadd(_cArray5,AllTrim(RETSALDO(_cArray2[_zz][1],"11")))
				Aadd(_cArray5,AllTrim(RETSALDO(_cArray2[_zz][1],"12")))				
				
				fWrite(_nHandle,'    <Cell ss:StyleID="s76"><Data ss:Type="Number">'+Strtran(Alltrim(Transform(Val(_cArray5[1]), "@E 999999999.99" )),",",".")+'</Data></Cell>'+ c_ent)
				fWrite(_nHandle,'    <Cell ss:StyleID="s76"><Data ss:Type="Number">'+Strtran(Alltrim(Transform(Val(_cArray5[2]), "@E 999999999.99" )),",",".")+'</Data></Cell>'+ c_ent)
				fWrite(_nHandle,'    <Cell ss:StyleID="s76"><Data ss:Type="Number">'+Strtran(Alltrim(Transform(Val(_cArray5[3]), "@E 999999999.99" )),",",".")+'</Data></Cell>'+ c_ent)
				fWrite(_nHandle,'    <Cell ss:StyleID="s76"><Data ss:Type="Number">'+Strtran(Alltrim(Transform(Val(_cArray5[4]), "@E 999999999.99" )),",",".")+'</Data></Cell>'+ c_ent)
				fWrite(_nHandle,'    <Cell ss:StyleID="s76"><Data ss:Type="Number">'+Strtran(Alltrim(Transform(Val(_cArray5[5]), "@E 999999999.99" )),",",".")+'</Data></Cell>'+ c_ent)
				fWrite(_nHandle,'    <Cell ss:StyleID="s76"><Data ss:Type="Number">'+Strtran(Alltrim(Transform(Val(_cArray5[6]), "@E 999999999.99" )),",",".")+'</Data></Cell>'+ c_ent)
				fWrite(_nHandle,'    <Cell ss:StyleID="s76"><Data ss:Type="Number">'+Strtran(Alltrim(Transform(Val(_cArray5[7]), "@E 999999999.99" )),",",".")+'</Data></Cell>'+ c_ent)
				fWrite(_nHandle,'    <Cell ss:StyleID="s76"><Data ss:Type="Number">'+Strtran(Alltrim(Transform(Val(_cArray5[8]), "@E 999999999.99" )),",",".")+'</Data></Cell>'+ c_ent)
				fWrite(_nHandle,'    <Cell ss:StyleID="s76"><Data ss:Type="Number">'+Strtran(Alltrim(Transform(Val(_cArray5[9]), "@E 999999999.99" )),",",".")+'</Data></Cell>'+ c_ent)
				fWrite(_nHandle,'    <Cell ss:StyleID="s76"><Data ss:Type="Number">'+Strtran(Alltrim(Transform(Val(_cArray5[10]), "@E 999999999.99" )),",",".")+'</Data></Cell>'+ c_ent)
				fWrite(_nHandle,'    <Cell ss:StyleID="s76"><Data ss:Type="Number">'+Strtran(Alltrim(Transform(Val(_cArray5[11]), "@E 999999999.99" )),",",".")+'</Data></Cell>'+ c_ent)
				fWrite(_nHandle,'    <Cell ss:StyleID="s76"><Data ss:Type="Number">'+Strtran(Alltrim(Transform(Val(_cArray5[12]), "@E 999999999.99" )),",",".")+'</Data></Cell>'+ c_ent)
				fWrite(_nHandle,'    <Cell ss:StyleID="s76"/>'+ c_ent)
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
		fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
		fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
		fWrite(_nHandle,'   </Row>'+ c_ent)                         
		
Next _cc

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////		

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

_cRetSoma := LINHACALC(_cArray9)
		
fWrite(_nHandle,'   <Row ss:AutoFitHeight="0">'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s63"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s77"><Data ss:Type="String">TOTAL GERAL</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s77"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s79" ss:Formula="='+_cRetSoma+'"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s79" ss:Formula="='+_cRetSoma+'"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s79" ss:Formula="='+_cRetSoma+'"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s79" ss:Formula="='+_cRetSoma+'"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s79" ss:Formula="='+_cRetSoma+'"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s79" ss:Formula="='+_cRetSoma+'"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s79" ss:Formula="='+_cRetSoma+'"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s79" ss:Formula="='+_cRetSoma+'"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s79" ss:Formula="='+_cRetSoma+'"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s79" ss:Formula="='+_cRetSoma+'"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s79" ss:Formula="='+_cRetSoma+'"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s79" ss:Formula="='+_cRetSoma+'"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s79" ss:Formula="=SUM(RC[-12]:RC[-1])"><Data ss:Type="Number">0</Data></Cell>'+ c_ent)
fWrite(_nHandle,'   </Row>'+ c_ent)
fWrite(_nHandle,'   <Row ss:AutoFitHeight="0">'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s63"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s75"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s75"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s71"/>'+ c_ent)
fWrite(_nHandle,'   </Row>'+ c_ent)
fWrite(_nHandle,'   <Row ss:AutoFitHeight="0">'+ c_ent)
fWrite(_nHandle,'    <Cell ss:Index="2" ss:StyleID="s70"/>'+ c_ent)
fWrite(_nHandle,'    <Cell ss:StyleID="s70"/>'+ c_ent)
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
fWrite(_nHandle,'     <ActiveRow>10</ActiveRow>'+ c_ent)
fWrite(_nHandle,'     <ActiveCol>1</ActiveCol>'+ c_ent)
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

RestArea(_cOIFR001)

Return()   

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³OIFR001   ºAutor  ³Leandro Ribeiro     º Data ³  19/06/14   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Função para retornar as contas atras da conta superior     º±±
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
Local _cArray3  := {}

//_cQuery3 := " SELECT DISTINCT AKD_CO, AKD_CC " + c_ent
_cQuery3 := " SELECT DISTINCT AKD_CO " + c_ent
_cQuery3 += " FROM "+RETSQLNAME("AKD")+" AKD "+ c_ent
_cQuery3 += " INNER JOIN "+RETSQLNAME("AK5")+" AK5 "+ c_ent
_cQuery3 += " ON AKD_FILIAL = AK5_FILIAL "+ c_ent 
_cQuery3 += " AND AKD_CO = AK5_CODIGO "+ c_ent
_cQuery3 += " WHERE "+ c_ent
_cQuery3 += " AKD_TPSALD = '"+AllTrim(MV_PAR01)+"' "+ c_ent
_cQuery3 += " AND AKD_CC BETWEEN '"+AllTrim(MV_PAR04)+"' AND '"+AllTrim(MV_PAR05)+"' "+ c_ent
_cQuery3 += " AND AKD_ITCTB BETWEEN '"+AllTrim(MV_PAR02)+"' AND '"+AllTrim(MV_PAR03)+"' "+ c_ent
//_cQuery3 += " AND AKD_OPER = '"+Alltrim(MV_PAR06)+"' "+ c_ent   
_cQuery3 += " AND AKD_OPER BETWEEN '"+AllTrim(MV_PAR06)+"' AND '"+AllTrim(MV_PAR07)+"' "+ c_ent
_cQuery3 += " AND AK5_COSUP = '"+_cContSup+"' "+ c_ent   
_cQuery3 += " AND AKD_DATA BETWEEN '"+DTOS(MV_PAR08)+"' AND '"+DTOS(MV_PAR09)+"' "+ c_ent         
_cQuery3 += " AND AKD_CO BETWEEN '"+MV_PAR10+"' AND '"+MV_PAR11+"' "+ c_ent  
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

(_cTrab3)->(dbCloseArea())

RestArea(_cArea3)

Return(_cArray3)            


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³OIFR001   ºAutor  ³Leandro Ribeiro     º Data ³  19/06/14   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ Função para retornar os saldos das contas                  º±±
±±º          ³                                                            º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP                                                        º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

//Static Function RETSALDO(_cCO,_cCC,_cData)
Static Function RETSALDO(_cCO,_cData)

Local _cArea4    := GetArea()
Local _cQuery4   := ""
Local _cTrab4    := GetNextAlias() 
Local _cRetSal   := ""                    
Local _cTpSaldo  := AllTrim(MV_PAR01) 
Local _cItemCont := _cCO
//Local _cCCusto   := _cCC
Local _cOperDe   := Alltrim(MV_PAR06)
Local _cOperAte  := Alltrim(MV_PAR07)
Local _cDataDe   := DTOS(MV_PAR08)
Local _cDataAte  := DTOS(MV_PAR09)
Local _xMes		 := VAL(_cData)
Local _zDataDe
Local _zDataAte
Local _cAno		 := SubStr(DTOS(MV_PAR08),1,4)   
Local _UltimoDia := Last_Day(STOD(_cAno+AllTrim(STRZERO(_xMes,2))+"01")) // Retorna o Ultimo dia do mês 

_zDataDe  := SUBSTR(_cDataAte,1,4)+AllTrim(STRZERO(_xMes,2))+"01"
_zDataAte := SUBSTR(_cDataDe,1,4)+AllTrim(STRZERO(_xMes,2))+AllTrim(STR(_UltimoDia))

If(_cDataDe <= _zDataDe)           
	_cDataDe := _zDataDe
Else
	_cDataDe := "99999999"
EndIf

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ//
//"Tipo de Saldo" 			MV_PAR01  //
//"Item Contabil de?" 		MV_PAR02  //
//"Item Contabil Ate?" 		MV_PAR03  //
//"Centro de Custo de?" 	MV_PAR04  //
//"Centro de Custo ate?" 	MV_PAR05  //
//"Oper. Orçamentaria de?" 	MV_PAR06  //
//"Oper. Orçamentaria Ate?" MV_PAR07  //
//"Data de" 				MV_PAR08  //
//"Data Ate" 				MV_PAR09  //
//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ// 


_cQuery4 := " SELECT AKD_CLVLR,"
Do Case
	Case(_cTpSaldo == 'PR')  
	_cQuery4 += " SUM(CASE WHEN AKD_TPSALD IN ('PR') THEN AKD_VALOR1 * (CASE WHEN AKD_TIPO='1' THEN 1 ELSE -1 END) ELSE 0 END) as VLTOTAL"+ c_ent
	Case(_cTpSaldo == 'EM')  
	_cQuery4 += " SUM(CASE WHEN AKD_TPSALD IN ('EM') THEN AKD_VALOR1 * (CASE WHEN AKD_TIPO='1' THEN 1 ELSE -1 END) ELSE 0 END) as VLTOTAL"+ c_ent
	Case(_cTpSaldo == 'RE')  
	_cQuery4 += " SUM(CASE WHEN AKD_TPSALD IN ('RE') THEN AKD_VALOR1 * (CASE WHEN AKD_TIPO='1' THEN 1 ELSE -1 END) ELSE 0 END) as VLTOTAL"+ c_ent
	Case(_cTpSaldo == 'OR')  
	_cQuery4 += " SUM(CASE WHEN AKD_TPSALD IN ('OR') THEN AKD_VALOR1 * (CASE WHEN AKD_TIPO='1' THEN 1 ELSE -1 END) ELSE 0 END) as VLTOTAL"+ c_ent
	Case(_cTpSaldo == 'CT')  
	_cQuery4 += " SUM(CASE WHEN AKD_TPSALD IN ('CT') THEN AKD_VALOR1 * (CASE WHEN AKD_TIPO='1' THEN 1 ELSE -1 END) ELSE 0 END) as VLTOTAL"+ c_ent
	Case(_cTpSaldo == 'PG')  
	_cQuery4 += " SUM(CASE WHEN AKD_TPSALD IN ('PG') THEN AKD_VALOR1 * (CASE WHEN AKD_TIPO='1' THEN 1 ELSE -1 END) ELSE 0 END) as VLTOTAL"+ c_ent
	Case(_cTpSaldo == 'PR+EM+RE')  
	_cQuery4 += " SUM(CASE WHEN AKD_TPSALD IN ('PR','EM','RE') THEN AKD_VALOR1 * (CASE WHEN AKD_TIPO='1' THEN 1 ELSE -1 END) ELSE 0 END) as VLTOTAL"
	EndCase
_cQuery4 += " FROM "+RETSQLNAME("AKD")+" AKD "+ c_ent
_cQuery4 += " WHERE AKD.AKD_FILIAL = '"+xFilial("AKD")+"'"+ c_ent /*"+xFilial("AKD")+"*/
//_cQuery4 += " AND AKD_ITCTB = '"+_cItemCont+"'"+ c_ent
_cQuery4 += " AND AKD_CO = '"+_cItemCont+"'"+ c_ent
_cQuery4 += " AND AKD_DATA  BETWEEN '"+_zDataDe+"' AND '"+_zDataAte+"'"+ c_ent // Leandro Ribeiro 05/01/2014
//_cQuery4 += " AND AKD_OPER = '"+_cOperOrc+"'"+ c_ent  
_cQuery4 += " AND AKD_OPER BETWEEN '"+_cOperDe+"' AND '"+_cOperAte+"'"+ c_ent
//_cQuery4 += " AND AKD_CC = '"+_cCCusto+"'"+ c_ent
_cQuery4 += " AND AKD_CC BETWEEN '"+MV_PAR04+"' AND '"+MV_PAR05+"'"+ c_ent
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
		//_cRetSal := Alltrim(Transform((_cTrab4)->VLTOTAL, "@E 999999999.99" ))
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
±±ºPrograma  ³OIFR001   ºAutor  ³Leandro Ribeiro     º Data ³  19/06/14   º±±
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
//Local _cCCusto   := _cCC
//Local _cOperOrc  := Alltrim(MV_PAR06)
Local _cOperDe   := Alltrim(MV_PAR06)
Local _cOperAte  := Alltrim(MV_PAR07)
Local _cDataDe   := DTOS(MV_PAR08)
Local _cDataAte  := DTOS(MV_PAR09)

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ//
//"Tipo de Saldo" 			MV_PAR01  //
//"Item Contabil de?" 		MV_PAR02  //
//"Item Contabil Ate?" 		MV_PAR03  //
//"Centro de Custo de?" 	MV_PAR04  //
//"Centro de Custo ate?" 	MV_PAR05  //
//"Operação Orçamentaria" 	MV_PAR06  //
//"Data de" 				MV_PAR07  //
//"Data Ate" 				MV_PAR08  //
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
_cQuery5 += " AND AKD_OPER BETWEEN '"+_cOperDe+"' AND '"+_cOperAte+"'"+ c_ent
//_cQuery5 += " AND AKD_CC = '"+_cCCusto+"'"+ c_ent
_cQuery5 += " AND AKD_CC BETWEEN '"+MV_PAR04+"' AND '"+MV_PAR05+"'"+ c_ent
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
±±ºPrograma  ³OIFR001   ºAutor  ³Leandro Ribeiro     º Data ³  20/06/14   º±±
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
±±ºPrograma  ³OIFR001   ºAutor  ³Leandro Ribeiro     º Data ³  28/05/14   º±±
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


