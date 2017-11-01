REM  *****  BASIC  *****

' --- Declaracao das funcoes --- '

Rem Caixa de seleção com confirmacao
Function ContinuaTestes(Mensagem)

	Rem Caixa de confirmação
	Dim ConfirmeBox As Integer
	ConfirmeBox = MB_YESNO + MB_DEFBUTTON2 + MB_ICONQUESTION

	Rem Executa o código e espera por resposta do usuario
	ContinuaTestes() = MsgBox (Mensagem, ConfirmeBox)

End Function

Rem Nivel do teste em questão
Function NivelTeste(nTeste, Nivel)

	Rem Objetos utilizados pelo sistema
	Dim FunctionAccess As Object
	
	Rem Carrega o Sistema
	FunctionAccess = createUnoService("com.sun.star.sheet.FunctionAccess")
	
	Rem Executa a formula
	Dim args As Object
	Dim divisao As Integer
	
	Rem Executa a formula
	args = Array(CDbl(nTeste / (2 ^ Nivel)), 1)
	divisao = FunctionAccess.callFunction("CEILING", args)
	
	Rem Calcula o resultado e retorna o valor
	NivelTeste() = ((divisao MOD 2) + 1)
	
End Function

Rem Checa se o objeto em questão é uma tabela
Function ChecaTabela(TabelaTestar)

	Rem Objeto sem motivo
	Dim DummyObj As Object

	Rem Checa se o objeto contem uma tabela de modelo
	On Error Goto TabelaIncorreta
		
		Rem Tenta acessar o modelo da tabela
		DummyObj = TabelaTestar.Model
		
		Rem Tabela Correta
		ChecaTabela() = 0
		Exit Function
		
	TabelaIncorreta:
		Rem Tabela Incorreta
		ChecaTabela() = 1
	
End Function

Rem Pega a pasta atual
Function PastaAtual()
	
	Rem Carrega as bibliotecas de pasta
	GlobalScope.BasicLibraries.loadLibrary("Tools")
	
	Rem Pega a pasta atual
    PastaAtual() = Tools.Strings.DirectoryNameoutofPath(ThisComponent.url, "/")

End Function

Rem Executa Uma musica
Sub ExecutaMusica(i_soundpath as string)
    Dim oPlayer1 As Object
    Dim sUrlSound As String
    Dim oSounMgr As Object
     
	Rem Checa se o gerenciador de musica não é nulo
    If Not IsNull(oSounMgr) then
    	S_Start_New()
        Exit Sub
    End If
    
    Rem Converte o caminho da musica 
    sUrlSound = ConvertToUrl(i_soundpath)

	Rem Checa se o arquivo existe e caso exista executa a musica
    If Not FileExists(sUrlSound) Then       
    	MsgBox(sUrlSound & " Nao Existe!", 16)
    Else
    
    	Rem Checa o sistema operacional do usuario
        If GetGuiType() = 1 Then
        	Rem Windows usa o DirectX
            oSounMgr = CreateUnoService("com.sun.star.media.Manager_DirectX")
        Else
        	Rem Linux Usa o GStreamer
            oSounMgr = CreateUnoService("com.sun.star.comp.media.Manager_GStreamer")
        End If
        
        Rem Checa se o sistema de execucao de musicas foi carregado
        If IsNull(oSounMgr) Then
           MsgBox("Player de musica nao carregado!", 16)
        Else
        
           	Rem Carrega a música
           	oPlayer1 = oSounMgr.createPlayer(sUrlSound)
           	
           	Rem Configura o player
           	oPlayer1.setMediaTime(0.0)
           	oPlayer1.setVolumeDB(-10)
           	oPlayer1.setPlayBackLoop(0)
           	oPlayer1.start(0)
           
			Rem Espera pela execução completa da música
           	While oPlayer1.isplaying()
       			doevents
           	WEnd
           
			Rem Limpa os objetos
           	oPlayer1 = Nothing
           	oSounMgr = Nothing
        End If
        
	End If
     
End Sub

' --- Declaracao das macros --- '

Rem Gemidao Open Office
Sub GemidaoOpenOffice()
	
	Rem Carrega a musica da pegadinha
	ExecutaMusica(PastaAtual() & "/apresentacao_topicos_mestrado.audio.wav")
	
	Rem Mensagem de pegadinha 
	MsgBox("Você acaba de cair no gemidão do OpenOffice!")
	
End Sub

Rem Completa os valores do sistema
Sub CompletaValores()

	Rem Carrega as variaveis
	GlobalScope.BasicLibraries.LoadLibrary("Tools")

	Rem Documento
	Dim Doc As Object
	
	Rem Página
	Dim Page As Object
	
	Rem Carrega o documento
	Doc = ThisComponent
	
	Rem Executa o código
	If ContinuaTestes("Deseja preencher a tabela?") = IDNO Then
		Exit Sub
	End If

	Rem Array de Niveis dos Fatores
	Dim FatorIS(2) As String
	Dim FatorIH(2) As String
	Dim FatorQM(2) As String
	Dim FatorPR(2) As String
	Dim FatorGE(2) As String
	Dim FatorDE(2) As String
	Dim FatorNC(2) As String
	Dim FatorTR(2) As String
	
	Rem Valores de cada nível de cada fator
	' ------------- '
	FatorIS(2) = "DC" 
	FatorIS(1) = "HV"
	' ------------- '
	FatorIH(2) = "HDC" 
	FatorIH(1) = "HHV"
	' ------------- '
	FatorQM(2) = "8GB" 
	FatorQM(1) = "16GB"
	' ------------- '
	FatorPR(2) = "i5" 
	FatorPR(1) = "i7"
	' ------------- '
	FatorGE(2) = "16" 
	FatorGE(1) = "32"
	' ------------- '
	FatorDE(2) = "160" 
	FatorDE(1) = "320"
	' ------------- '
	FatorNC(2) = "32" 
	FatorNC(1) = "320"
	' ------------- '
	FatorTR(2) = "1s" 
	FatorTR(1) = "10s"
	' -------------
	
	Rem Para cada página escrever o conteudo de cada coluna
	Dim i As Integer ' Iterador de páginas
	Dim n As Integer ' Iterador de testes
	
	Rem Inicia o contador de testes
	n = 1
	
	Rem Inicia a composição
	For i = 32 To 57
		' MsgBox(i & " Hehe") ; Caixas de mensagem são bloqueantes e fecham se o usuario desejar
		' Print("Página " & i) ; Imprimir bloqueia e dá a opção ao usuario de parar a execução do código
		
		Rem Confirmação de Execução
		'Print("Populando pagina " & i)
		
		Rem Carrega a página
		Page = Doc.DrawPages(i)
		
		Rem Variavel da tabela do slide
		Dim TabelaPage As Object
		
		Rem Procura pela Tabela dentre os objetos da página
		Dim k As Integer
		For k = 0 To Page.count() - 1
		
			Rem Pega um objeto da página
			TabelaPage = Page.getByIndex(k)
			
			Rem Se ele conter a propriedade Model então temos uma tabela
			If ChecaTabela(TabelaPage) = 0 Then
				Exit For
			End If
					
		Next k
		
		Rem Carrega o modelo da tabela
		Dim ModeloTabela As Object
		ModeloTabela = TabelaPage.Model
		
		Rem Popula a página de acordo com as variaveis
		Dim j As Integer
		For j = 2 To 11
			
			Rem Se chegarmos a última entrada então paramos a execução
			If n > 256 Then
				Exit For
			End If
			
			Rem Testes a serem efetuados
			Dim Teste As Object
			
			Rem Variavel de Nivel
			Dim Nivel As Integer
			
			' --- Executa os testes --- '
			
			Rem TesteIS
			Nivel = NivelTeste(n, 7)
			Teste = ModeloTabela.getCellByPosition(1, j)
			Teste.setString(FatorIS(Nivel))
			TesteCursor = Teste.createTextCursor()
         	TesteCursor.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER
			
			Rem TesteIH
			Nivel = NivelTeste(n, 6)
			Teste = ModeloTabela.getCellByPosition(2, j)
			Teste.setString(FatorIH(Nivel))
			TesteCursor = Teste.createTextCursor()
         	TesteCursor.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER
			
			Rem TesteQM
			Nivel = NivelTeste(n, 5)
			Teste = ModeloTabela.getCellByPosition(3, j)
			Teste.setString(FatorQM(Nivel))
			TesteCursor = Teste.createTextCursor()
         	TesteCursor.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER
			
			Rem TestePR
			Nivel = NivelTeste(n, 4)
			Teste = ModeloTabela.getCellByPosition(4, j)
			Teste.setString(FatorPR(Nivel))
			TesteCursor = Teste.createTextCursor()
         	TesteCursor.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER
			
			Rem TesteGE
			Nivel = NivelTeste(n, 3)
			Teste = ModeloTabela.getCellByPosition(5, j)
			Teste.setString(FatorGE(Nivel))
			TesteCursor = Teste.createTextCursor()
         	TesteCursor.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER
			
			Rem TesteDE
			Nivel = NivelTeste(n, 2)
			Teste = ModeloTabela.getCellByPosition(6, j)
			Teste.setString(FatorDE(Nivel))
			TesteCursor = Teste.createTextCursor()
         	TesteCursor.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER
			
			Rem TesteNC
			Nivel = NivelTeste(n, 1)
			Teste = ModeloTabela.getCellByPosition(7, j)
			Teste.setString(FatorNC(Nivel))
			TesteCursor = Teste.createTextCursor()
         	TesteCursor.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER
			
			
			Rem TesteTR
			Nivel = NivelTeste(n, 0)
			Teste = ModeloTabela.getCellByPosition(8, j)
			Teste.setString(FatorTR(Nivel))
			TesteCursor = Teste.createTextCursor()
         	TesteCursor.paraAdjust = com.sun.star.style.ParagraphAdjust.CENTER
			
			Rem Incrementa n
			n = n + 1
		
		Next j
		
	Next i
	
	Rem Fim da execução
	MsgBox("Tabelas preenchidas!")

End Sub


Rem Limpa todos os valores
Sub LimpaValores()

	Rem Carrega as variaveis
	GlobalScope.BasicLibraries.LoadLibrary("Tools")

	Rem Documento
	Dim Doc As Object
	
	Rem Página
	Dim Page As Object
	
	Rem Carrega o documento
	Doc = ThisComponent
	
	Rem Executa o código
	If ContinuaTestes("Deseja limpar a tabela?") = IDNO Then
		Exit Sub
	End If
	
	Rem Para cada página escrever o conteudo de cada coluna
	Dim i As Integer ' Iterador de páginas
	Dim n As Integer ' Iterador de testes
	
	Rem Inicia o contador de testes
	n = 1

	Rem Inicia a composição
	For i = 32 To 57
		' MsgBox(i & " Hehe") ; Caixas de mensagem são bloqueantes e fecham se o usuario desejar
		' Print("Página " & i) ; Imprimir bloqueia e dá a opção ao usuario de parar a execução do código
		
		Rem Confirmação de Execução
		'Print("Limpando pagina " & i)
		
		Rem Carrega a página
		Page = Doc.DrawPages(i)
		
		Rem Variavel da tabela do slide
		Dim TabelaPage As Object
		
		Rem Procura pela Tabela dentre os objetos da página
		Dim k As Integer
		For k = 0 To Page.count() - 1
		
			Rem Pega um objeto da página
			TabelaPage = Page.getByIndex(k)
			
			Rem Se ele conter a propriedade Model então temos uma tabela
			If ChecaTabela(TabelaPage) = 0 Then
				Exit For
			End If
			
		Next k
		
		Rem Carrega o modelo da tabela
		Dim ModeloTabela As Object
		ModeloTabela = TabelaPage.Model
		
		Rem Popula a página de acordo com as variaveis
		Dim j As Integer
		For j = 2 To 11
			
			Rem Se chegarmos a última entrada então paramos a execução
			If n > 256 Then
				Exit For
			End If
			
			Rem Testes a serem efetuados
			Dim Teste As Object
			
			' --- Executa os testes --- '
			
			Rem TesteIS
			Teste = ModeloTabela.getCellByPosition(1, j)
			Teste.setString("")
			
			Rem TesteIH
			Teste = ModeloTabela.getCellByPosition(2, j)
			Teste.setString("")
			
			Rem TesteQM
			Teste = ModeloTabela.getCellByPosition(3, j)
			Teste.setString("")
			
			Rem TestePR
			Teste = ModeloTabela.getCellByPosition(4, j)
			Teste.setString("")
			
			Rem TesteGE
			Teste = ModeloTabela.getCellByPosition(5, j)
			Teste.setString("")
			
			Rem TesteDE
			Teste = ModeloTabela.getCellByPosition(6, j)
			Teste.setString("")
			
			Rem TesteNC
			Teste = ModeloTabela.getCellByPosition(7, j)
			Teste.setString("")
			
			Rem TesteTR
			Teste = ModeloTabela.getCellByPosition(8, j)
			Teste.setString("")
			
			Rem Incrementa n
			n = n + 1
		
		Next j
		
	Next i

	Rem Fim da execução
	MsgBox("Tabelas limpas!")

End Sub
