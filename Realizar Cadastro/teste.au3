
#include<Excel.au3>
;#include<variaveis.au3>
#include<manutencao_clientes.au3>

HotKeySet("w", "sair")

Func ler_valores_usuario() ; Vai abrir a planilha e ler os valores

	Local $oExcel = _Excel_Open() ; Abrir a planilha
	Local $oWorkbook = _Excel_BookOpen($oExcel, @ScriptDir & "\excel\user1.xls") ; Caminho do arquivo
	Global $sResultUser = _Excel_RangeRead($oWorkbook, Default, "A" & $user) ; Vai pegar o valor da coluna A na posição que estiver a variável $user (No caso 2)
	Global $sResultPass = _Excel_RangeRead($oWorkbook, Default, "B" & $password) ; Vai pegar o valor da coluna B na posição que estiver a variável $password (No caso 2)
	_Excel_Close($oExcel) ; fechar planilha

EndFunc

Func logar_retaguarda()

	Run("C:\Facil\Retaguarda\Retaguarda.exe"); Abre o Retaguarda

	WinWaitActive($r)
	WinActivate($r)
	ControlSend($r, "", "Edit1", $sResultUser) ; Muda o que tiver no campo pra o que eu definir (Linha 17)
	ControlSend($r,"","TEdit1", $sResultPass); (Linha 18)
	ControlClick($r, "" , "TJvXPButton1", "left", 1) ;Login
EndFunc



Func verificar_login()

	ler_valores_usuario() ; Primeiro vai abrir e ler os valores que eu coloquei no Excel
	logar_retaguarda()
	WinWaitActive($versao, "", 5) ; Vai esperar 5 segundos (Se nesse tempo o Retaguarda não for aberto significa que deu erro no login então ele executa a linha 52)
EndFunc

Func cadastrar_cliente_pessoa_fisica()

	ControlClick($versao,"", "TPanel2", "left",1) ;Clica na opção clientes
	WinWaitActive($tela_consulta_clientes)
	ControlClick($tela_consulta_clientes, "", "TJvXPButton1", "left",1) ;Inserir novo cliente
	WinWaitActive("Confirmação!")
	ControlClick("Confirmação!", "", "TJvXPButton2", "left",1) ;Confirmação pessoa Jurídica
	WinWaitActive($tela_cadastro_clientes)
EndFunc

verificar_login()
WinWaitActive($versao)
cadastrar_cliente_pessoa_fisica()
inserir_manutencao_cliente()
Sleep(1000)


Func fechar()
	WinWaitActive($tela_cadastro_clientes)
	Sleep(1000)
	WinClose($tela_cadastro_clientes)
	WinWaitActive("Confirmação!")
	ControlClick("Confirmação!", "", "TJvXPButton2", "left", 1)
EndFunc



Func sair()
	Exit
EndFunc

