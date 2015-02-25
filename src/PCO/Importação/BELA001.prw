#include "protheus.ch"
#INCLUDE "rwmake.ch"
#INCLUDE "topconn.ch"
#DEFINE   c_ent      CHR(13)+CHR(10)

/*/
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  ณBELA001	  Autor ณ Marcelo Amaral     บ Data ณ  02/02/15   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDescricao ณ Rotina para importa็ao de arquivo csv.           	      บฑฑ
ฑฑบ          ณ Importa็ใo de itens da planilha or็amentaria.		 	  บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ AP6 IDE                                  			      บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
/*/

User Function BELA001()


//ฺฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฟ
//ณ Declaracao de Variaveis                                             ณ
//ภฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤู

Private oLeTxt
Private cArqCsv   := ""
Private aErros := {}

dbSelectArea("AK2")
dbSetOrder(1)      

//ฺฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฟ
//ณ Montagem da tela de processamento.                                  ณ
//ภฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤู

@ 163,1 TO 360,400 DIALOG oLeTxt TITLE OemToAnsi("Leitura de Arquivo CSV") 
@ 02,10 TO 080,190
@ 010,018 Say " Este programa ira ler o conteudo de um arquivo csv, conforme"
@ 018,018 Say " os parametros definidos pelo usuario. 						"
@ 026,018 Say " Para importa็ใo de itens da planilha or็amentaria.          "

@ 83,103 BMPBUTTON TYPE 14 ACTION (cArqCsv:=cGetFile("Arquivo Texto (*.csv) | *.csv",OemToAnsi("Selecione Diretorio"),,"",.F.,GETF_LOCALHARD+GETF_NETWORKDRIVE,.F.))
@ 83,133 BMPBUTTON TYPE 01 ACTION OkLeTxt()
@ 83,163 BMPBUTTON TYPE 02 ACTION Close(oLeTxt)

Activate Dialog oLeTxt Centered

Return      

/*/
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบFuno    ณ OKLETXT  บ Autor ณ AP6 IDE            บ Data ณ  06/12/11   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDescrio ณ Funcao chamada pelo botao OK na tela inicial de processamenบฑฑ
ฑฑบ          ณ to. Executa a leitura do arquivo texto.                    บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ Programa principal                                         บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
/*/

Static Function OkLeTxt

//ฺฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฟ
//ณ Abertura do arquivo texto                                           ณ
//ภฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤู


nHdl := FT_FUSE(cArqCsv)

If nHdl == -1
    MsgAlert("O arquivo de nome "+cArqCsv+" nใo pode ser aberto! Verifique os parametros.","Aten็ใo!")
    Return
Endif

//ฺฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฟ
//ณ Inicializa a regua de processamento                                 ณ
//ภฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤู

Processa({|| RunCont(cArqCsv) },"Processando...")
Return  

/*/
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบFuno    ณ RUNCONT  บ Autor ณ AP5 IDE            บ Data ณ  06/12/11   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDescrio ณ Funcao auxiliar chamada pela PROCESSA.  A funcao PROCESSA  บฑฑ
ฑฑบ          ณ monta a janela com a regua de processamento.               บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ Programa principal                                         บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
/*/

Static Function RunCont(cArqCsv)

Local _aArea1   := GetArea() 
Local nTamFile, nTamLin, cBuffer, nBtLidos
Local _cInfo    := {}
Local xPos      := "" 
Local _cDados	:= {}
Local _cMes		:= {}
Local _xRet		
Local _cFile    := "LOG " + DToS( Date() ) + ".txt"
Local _cPath    := "C:\LogSiga"
Local _nArqRef 	:= 0   
Local _cCGC		:= ""
Local _cMoeda	:= ""
Local _ee 		:= 1 
Local _aCabSF1	:= {}
Local _aItensSD1:= {}
Local aBaixa	:= {}
Local _cNumNot  := ""
Local cConta	:= 0
Local _xxIRRF	:= 0
Local _xxCSLL	:= 0 
Local _xxCofins := 0
Local _xxISS	:= 0
Local _xxPIS	:= 0  
Local _cNumDoc	:= ""
Local _cSerDoc	:= ""
Local _cMes		:= {}

Private lMSErroAuto	   :=.F.
Private lMsHelpAuto    :=.F.
Private cArquiv   := cArqCsv
Private aItens 	:= {}

//ฺฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฟ
//ณMonta o Nome do Arquivoณ
//ภฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤู
_cArquivo := Alltrim( _cPath ) + Iif( Right( Alltrim( _cPath ), 1 ) == "\", "", "\" ) + Alltrim( _cFile )
If File( _cArquivo )
	FErase( _cArquivo )
End If

//ฺฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฟ
//ณCria o Arquivo Fisicoณ
//ภฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤู
_nArqRef := MsFCreate( _cArquivo )

MakeDir("C:\LogSiga")

//ษอออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออป
//บ Lay-Out do arquivo Texto gerado:                                บ
//ฬออออออออออออออออัออออออออัอออออออออออออออออออออออออออออออออออออออน
//บCampo           ณ Inicio ณ Tamanho                               บ
//วฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤลฤฤฤฤฤฤฤฤลฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤถ
//	N๚mero da nota			aItens[1]                              //
//	S้rie					aItens[2]                              //
//	Data de Emissใo			aItens[3]                              //
//	Fornecedor				aItens[4]                              //
//	Esp้cie					aItens[5]                              //
//	UF origem				aItens[6]                              //
//	Produto					aItens[7]                              //
//	Quantidade				aItens[8]                              //
//	Valor Unitแrio			aItens[9]                              //
//	Tipo de entrada (TES)	aItens[10]                             //
//	Valor da glosa			aItens[11]                             //
//	Gera DIRF				aItens[12]                             //
//	INSS EMP				aItens[13]                             //
//	INSS CRED				aItens[14]                             //
//	IRRF					aItens[15]                             //
//	PIS						aItens[16]                             //
//	COFINS					aItens[17]                             //
//	CSLL					aItens[18]                             //
//	ISS						aItens[19]                             //
//	Condi็ใo de pagamento	aItens[20]                             //
//	Natureza				aItens[21]                             //
//	Data Vencimento			aItens[22]                             //
//	Data Pagamento			aItens[23]                             //
//ศออออออออออออออออฯออออออออฯอออออออออออออออออออออออออออออออออออออออผ

*-------------------*
*Fecha tela inicial *
*-------------------*
Close(oLeTxt)

*-----------------------*
* Abre o Arquivo Texto  *
*-----------------------*
FT_FUSE(cArquiv)

*-----------------------------------------------------------------------------------------*
* Vai para o Inicio do Arquivo e Define o numero de Linhas para a Barra de Processamento. *
*-----------------------------------------------------------------------------------------*
FT_FGOTOP()
ProcRegua(FT_FLASTREC())
_cTotal := FT_FLASTREC()


*-----------------------------*
* Leitura da linha do arquivo *
*-----------------------------*
cBuffer := FT_FREADLN()
FT_FSKIP()  

nLinhaAtu := 0

While !FT_FEOF()

    //ฺฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฟ
    //ณ Incrementa a regua                                                  ณ
    //ภฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤู
	cConta := cConta + 1
    //IncProc("Importando Arquivo Aguarde...")
    
    nLinhaAtu += 1
     
    IncProc("Importando Arquivo "+ Alltrim(str(cConta)) +"/"+ Alltrim(str(_cTotal - 1)) +" Aguarde...")
    ProcRegua(_cTotal)
	*----------------------------------------------------------------*
	* Faz a Leitura da Linha do Arquivo e atribui a Variavel cBuffer *
	*----------------------------------------------------------------*
	cBuffer := FT_FREADLN()
	
	*-------------------------------------------------------------*
	* Se ja passou por todos os registros da planilha "CSV" sai do While. *
	*-------------------------------------------------------------*
	If Empty(cBuffer)
		Exit
	Endif
	
	*---------------------------------------------*
	* Retorna posicao em que foi encontrado o ";" *
	*---------------------------------------------*
	xPos := AT(";",cBuffer)
	
	*----------------------------------------------------------------------------------------*
	* Adiciona as informacoes no vetor ate o ";" e retira o conteudo inserido no vetor da linha cBuffer *
	*----------------------------------------------------------------------------------------*
	While xPos <> 0
		_cInfo := Alltrim(SubStr(cBuffer , 1, xPos-1 ))
		aAdd( aItens , _cInfo )
		cBuffer:= SubStr(cBuffer , xPos+1, Len(cBuffer)-xPos)
		xPos := AT(";",cBuffer)
	Enddo
	*---------------------------------------------------------------------------------------*
	* Se ainda tiver informacao no cBuffer, adiciona a ultima parte do arquivo depois do ";" no vetor *
	*---------------------------------------------------------------------------------------*
	if Len(cBuffer)>0
		_cInfo := Alltrim(SubStr(cBuffer , 1, Len(cBuffer) ))
		aAdd( aItens , _cInfo )
	endif
	aAdd( _cDados , aItens )
	
	Aadd(_cMes, aItens[20])
	Aadd(_cMes, aItens[21])
	Aadd(_cMes, aItens[22])
	Aadd(_cMes, aItens[23])
	Aadd(_cMes, aItens[24])
	Aadd(_cMes, aItens[25])
	Aadd(_cMes, aItens[26])
	Aadd(_cMes, aItens[27])
	Aadd(_cMes, aItens[28])
	Aadd(_cMes, aItens[29])
	Aadd(_cMes, aItens[30])
	Aadd(_cMes, aItens[31])
//	Aadd(_cMes, aItens[32])
	 
	if !U_BELA001A()
		aItens := {}
		_cMes  := {}   
		FT_FSKIP() 
		Loop
	Endif   
	
    //IncProc("Importando Arquivo Aguarde...")
    IncProc("Importando Arquivo "+ Alltrim(str(cConta)) +"/"+ Alltrim(str(_cTotal - 1)) +" Aguarde...") 
	ProcRegua(RecCount())
    //ฺฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฟ
    //ณ Grava os campos obtendo os valores da linha lida do arquivo texto.  ณ
    //ภฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤู
  
  	For _cc := 1 to Len(_cMes)
  
    	DbSelectArea("AK2") 
	   	DbSetorder(1)
	   	_xRet := !(DbSeek(Xfilial("AK2")+padr(aItens[2],tamsx3("AK2_ORCAME")[1])+padr(aItens[3],tamsx3("AK2_VERSAO")[1])+padr(aItens[4],tamsx3("AK2_CO")[1])+aItens[32]+AllTrim(STRZERO(_cc,2))+"01"+AllTrim(Strzero(Val(aItens[1]),4))))
	       
		RecLock("AK2",_xRet) 
	                                     
        AK2->AK2_FILIAL 	:= xFilial("AK2")                  
  	    AK2->AK2_ID         := AllTrim(Strzero(Val(aItens[1]),4))         
       	AK2->AK2_ORCAME     := aItens[2]
       	AK2->AK2_VERSAO     := AllTrim(Strzero(Val(aItens[3]),4))
        AK2->AK2_CO         := aItens[4]
        AK2->AK2_PERIOD     := STOD(aItens[32]+AllTrim(STRZERO(_cc,2))+"01")//aItens[5]
        AK2->AK2_CC         := aItens[12]
        AK2->AK2_ITCTB      := aItens[16]
        AK2->AK2_CLVLR      := aItens[14]
        AK2->AK2_CLASSE     := aItens[6]
        AK2->AK2_VALOR      := Val(StrTran(_cMes[_cc],",",".")) 
        AK2->AK2_DESCRI     := aItens[8]
        AK2->AK2_OPER       := aItens[10]
        AK2->AK2_MOEDA 		:= 1
        //AK2->AK2_CHAVE
        AK2->AK2_DATAF      := Stod(Substr(DtoS(ddatabase),1,4)+AllTrim(STRZERO(_cc,2))+"01")//aItens[5]// usar fun็ใo para trazer o ultimo dia do mes
        AK2->AK2_DATAI      := Stod(Substr(DtoS(ddatabase),1,4)+AllTrim(STRZERO(_cc,2))+"01")//aItens[5]
        AK2->AK2_UNIORC     := aItens[9]
	        
		MsUnlock()  
	     
	 	PcoIniLan("000252")//PROCESSO "000252 = IMPORTAวรO DE PLANILHA"
		PcoDetLan("000252","01","BELA001")
    	PcoFinLan("000252")	//     
	   
	Next _cc 
	
			
 	//ฺฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฟ
	//ณ Leitura da proxima linha do arquivo texto.                          ณ
	//ภฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤู
	
	aItens := {}
	_cMes  := {}   
	
	FT_FSKIP() 
	  
EndDo

FT_FUSE()   

If !empty(aErros)
	Aviso("ATENวรO", "Foram encontradas inconsist๊ncias na importa็ใo."+chr(13)+chr(10)+"Serแ gravada uma Planilha com as Inconsist๊ncias.",{"Ok"})
	GravaLog()
Else
	Aviso("ATENวรO","Importa็ใo finalizada com sucesso!",{"Ok"})
EndIf

//Aviso("A V I S O !","Planilha Importada!",{"OK"})

//ฺฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฟ
//ณ O arquivo texto deve ser fechado, bem como o dialogo criado na fun- ณ
//ณ cao anterior.                                                       ณ
//ภฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤู

fClose(nHdl)
Close(oLeTxt)

RestArea(_aArea1)
	
Return     

User Function BELA001A()
Local lRet := .T.
Local aLogAux := {}

if Alltrim(aItens[2]) == "" .or. AllTrim(Strzero(Val(aItens[3]),4)) == ""
	//aadd(aLogAux,"Planilha Or็amentแria nใo Informada!"+c_ent)
	//lRet := .F.
else
	dbSelectArea("AK1")
	AK1->(dbSetOrder(1))
	if !AK1->(dbSeek(xFilial("AK1")+Padr(aItens[2],TamSX3("AK2_ORCAME")[1])+AllTrim(Strzero(Val(aItens[3]),4)) ))
		aadd(aLogAux,"Planilha Or็amentแria nใo Cadastrada!"+c_ent)
		lRet := .F.
	endif
endif

if Alltrim(aItens[4]) == ""
	//aadd(aLogAux,"Conta Or็amentแria nใo Informada!"+c_ent)
	//lRet := .F.
else
	dbSelectArea("AK5")
	AK5->(dbSetOrder(1))
	if !AK5->(dbSeek(xFilial("AK5")+aItens[4] ))
		aadd(aLogAux,"Conta Or็amentแria nใo Cadastrada!"+c_ent)
		lRet := .F.
	endif
endif

if Alltrim(aItens[6]) == ""
	//aadd(aLogAux,"Classe Or็amentแria nใo Informada!"+c_ent)
	//lRet := .F.
else
	dbSelectArea("AK6")
	AK6->(dbSetOrder(1))
	if !AK6->(dbSeek(xFilial("AK6")+aItens[6] ))
		aadd(aLogAux,"Classe Or็amentแria nใo Cadastrada!"+c_ent)
		lRet := .F.
	endif
endif

if Alltrim(aItens[9]) == ""
	//aadd(aLogAux,"Unidade Or็amentแria nใo Informada!"+c_ent)
	//lRet := .F.
else
	dbSelectArea("AMF")
	AMF->(dbSetOrder(1))
	if !AMF->(dbSeek(xFilial("AMF")+aItens[9] ))
		aadd(aLogAux,"Unidade Or็amentแria nใo Cadastrada!"+c_ent)
		lRet := .F.
	endif
endif

if Alltrim(aItens[10]) == ""
	//aadd(aLogAux,"Opera็ใo Or็amentแria nใo Informada!"+c_ent)
	//lRet := .F.
else
	dbSelectArea("AKF")
	AKF->(dbSetOrder(1))
	if !AKF->(dbSeek(xFilial("AKF")+aItens[10] ))
		aadd(aLogAux,"Opera็ใo Or็amentแria nใo Cadastrada!"+c_ent)
		lRet := .F.
	endif
endif

if Alltrim(aItens[12]) == ""
	//aadd(aLogAux,"Centro de Custos nใo Informado!"+c_ent)
	//lRet := .F.
else
	dbSelectArea("CTT")
	CTT->(dbSetOrder(1))
	if !CTT->(dbSeek(xFilial("CTT")+aItens[12] ))
		aadd(aLogAux,"Centro de Custos nใo Cadastrado!"+c_ent)
		lRet := .F.
	endif
endif

if Alltrim(aItens[16]) == ""
	//aadd(aLogAux,"Item Contแbil nใo Informado!"+c_ent)
	//lRet := .F.
else
	dbSelectArea("CTD")
	CTD->(dbSetOrder(1))
	if !CTD->(dbSeek(xFilial("CTD")+aItens[16] ))
		aadd(aLogAux,"Item Contแbil nใo Cadastrado!"+c_ent)
		lRet := .F.
	endif
endif

if Alltrim(aItens[14]) == ""
	//aadd(aLogAux,"Classe de Valor nใo Informada!"+c_ent)
	//lRet := .F.
else
	dbSelectArea("CTH")
	CTH->(dbSetOrder(1))
	if !CTH->(dbSeek(xFilial("CTH")+aItens[14] ))
		aadd(aLogAux,"Classe de Valor nใo Cadastrada!"+c_ent)
		lRet := .F.
	endif
endif

if Len(aLogAux) > 0
	
	aAdd(aErros, "Linha: " + alltrim(str(nLinhaAtu)) + chr(13) + chr(10) )
	
    For nCont := 1 to Len(aLogAux)
    	
    	aadd(aErros,aLogAux[nCont])
    	
    Next

	aAdd(aErros, Replicate("-",50) + chr(13) + chr(10) )

endif

Return lRet

Static Function GravaLog()
Local nArq

cArqLog := strTran(upper(cArquiv),".CSV","") + "_LOG_" + Dtos(dDataBase) + "_" + StrTran(Time(),":","") + ".LOG"

nArq := FCreate(cArqLog)

If nArq > 0
	ProcRegua(len(aErros))
	For i := 1 to len(aErros)
		incProc("Gravando o Log de Erros . . . ")
		FWrite(nArq, aErros[i])
	Next i
	FClose(nArq)
	If ApOleClient("MSExcel") // testa a integra็ใo com o excel.
		oExcel := MSExcel():New()
		oExcel:WorkBooks:Open(cArqLog)
		oExcel:SetVisible(.T.)
		oExcel:Destroy()
	endif
	MsgInfo("Gera็ใo da Planilha de Inconsist๊ncias "+cArqLog+" Finalizada!")
Else 
	MsgAlert("Erro na gera็ใo da Planilha de Inconsist๊ncias.")
EndIf

Return
