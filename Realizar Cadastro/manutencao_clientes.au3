#include<variaveis.au3>
#include<Excel.au3>

Func inserir_manutencao_cliente() ; Vai abrir a planilha e ler os valores


	Local $oExcel = _Excel_Open() ; Abrir a planilha
	Local $oWorkbook = _Excel_BookOpen($oExcel, @ScriptDir & "\excel\user.xls") ; Caminho do arquivo
	Global $sResultNome_cliente = _Excel_RangeRead($oWorkbook, Default, "A" & $nome_cliente) ; Vai pegar o valor da coluna A na posição que estiver a variável $user (No caso 2)
	Global $sResultData_nascimento = _Excel_RangeRead($oWorkbook, Default, "B" & $data_nascimento)
	Global $sResultCpf_cliente = _Excel_RangeRead($oWorkbook, Default, "C" & $cpf_cliente)
	Global $sResultIdentidade_cliente = _Excel_RangeRead($oWorkbook, Default, "D" & $identidade_cliente)
	Global $sResultCartao_identidade_cliente = _Excel_RangeRead($oWorkbook, Default, "E" & $cartao_identidade_cliente)
	Global $sResultContato_representante_cliente = _Excel_RangeRead($oWorkbook, Default, "F" & $contato_representante_cliente)
	Global $sResultCep_residencial_cliente = _Excel_RangeRead($oWorkbook, Default, "G" & $cep_residencial_cliente)
	Global $sResultEndereco_residencial_cliente = _Excel_RangeRead($oWorkbook, Default, "H" & $endereco_residencial_cliente)
	Global $sResultNumero_residencial_cliente= _Excel_RangeRead($oWorkbook, Default, "I" & $numero_residencial_cliente)
	Global $sResultComplemento_residencial_cliente = _Excel_RangeRead($oWorkbook, Default, "J" & $complemento_residencial_cliente)
	Global $sResultBairro_residencial_cliente = _Excel_RangeRead($oWorkbook, Default, "K" & $bairro_residencial_cliente)
	Global $sResultReferencia_residencial_cliente = _Excel_RangeRead($oWorkbook, Default, "L" & $referencia_residencial_cliente)
	Global $sResultTelefone_fixo_cliente = _Excel_RangeRead($oWorkbook, Default, "M" & $telefone_fixo_cliente)
	Global $sResultCelular_cliente = _Excel_RangeRead($oWorkbook, Default, "N" & $celular_cliente)
	Global $sResultDdd_telefone2_cliente = _Excel_RangeRead($oWorkbook, Default, "O" & $ddd_telefone2_cliente)
	Global $sResultTelefone2_cliente = _Excel_RangeRead($oWorkbook, Default, "P" & $telefone2_cliente)
	Global $sResultEmail_cliente = _Excel_RangeRead($oWorkbook, Default, "Q" & $email_cliente)
	Global $sResultContato_rede_social_cliente = _Excel_RangeRead($oWorkbook, Default, "R" & $contato_rede_social_cliente)
	Global $sResultFacebook_cliente = _Excel_RangeRead($oWorkbook, Default, "S" & $facebook_cliente)
	Global $sResultHomepage_cliente = _Excel_RangeRead($oWorkbook, Default, "T" & $homepage_cliente)
	Global $sResultCaixa_postal_cliente = _Excel_RangeRead($oWorkbook, Default, "U" & $caixa_postal_cliente)
	Global $sResultNome_pai_cliente = _Excel_RangeRead($oWorkbook, Default, "V" & $nome_pai_cliente)
	Global $sResultNome_mae_cliente = _Excel_RangeRead($oWorkbook, Default, "W" & $nome_mae_cliente)
	Global $sResultReferencia1_cliente = _Excel_RangeRead($oWorkbook, Default, "X" & $referencia1_cliente)
	Global $sResultReferencia2_cliente = _Excel_RangeRead($oWorkbook, Default, "Y" & $referencia2_cliente)
	Global $sResultReferencia3_cliente = _Excel_RangeRead($oWorkbook, Default, "Z" & $referencia3_cliente)
	Global $sResultInscricao_municipal_cliente = _Excel_RangeRead($oWorkbook, Default, "AA" & $inscricao_municipal_cliente)
	Global $sResultCampo_adicional1_cliente = _Excel_RangeRead($oWorkbook, Default, "AB" & $campo_adicional1_cliente)
	Global $sResultCampo_adicional2_cliente = _Excel_RangeRead($oWorkbook, Default, "AC" & $campo_adicional2_cliente)
	Global $sResultCampo_adicional3_cliente= _Excel_RangeRead($oWorkbook, Default, "AD" & $campo_adicional3_cliente)
	Global $sResultCampo_adicional4_cliente = _Excel_RangeRead($oWorkbook, Default, "AE" & $campo_adicional4_cliente)
	Global $sResultSexo_cliente = _Excel_RangeRead($oWorkbook, Default, "AF" & $sexo_cliente)
	Global $sResultObservacoes_cliente = _Excel_RangeRead($oWorkbook, Default, "AG" & $observacoes_cliente)
	Global $sResultLimite_credito_cliente = _Excel_RangeRead($oWorkbook, Default, "AH" & $limite_credito_cliente)
	;Global $sResultConvenio_numero = _Excel_RangeRead($oWorkbook, Default, "AI" & $convenio_numero)
	Global $sResultPreco_cliente = _Excel_RangeRead($oWorkbook, Default, "AJ" & $preco_cliente)
	Global $sResultVendedor_padrao_cliente = _Excel_RangeRead($oWorkbook, Default, "AK" & $vendedor_padrao_cliente)
	Global $sResultSituacao_cliente  = _Excel_RangeRead($oWorkbook, Default, "AL" & $situacao_cliente)
	Global $sResultTipo_cliente = _Excel_RangeRead($oWorkbook, Default, "AM" & $tipo_cliente)
	Global $sResultPrograma_fidelidade_cliente = _Excel_RangeRead($oWorkbook, Default, "AN" & $programa_fidelidade_cliente)
	Global $sResultBloquear_inadimplente_cliente = _Excel_RangeRead($oWorkbook, Default, "AO" & $bloquear_inadimplente_cliente)
	Global $sResultSetor_cliente = _Excel_RangeRead($oWorkbook, Default, "AP" & $setor_cliente)
	Global $sResultDdd_celular_cliente = _Excel_RangeRead($oWorkbook, Default, "AQ" & $ddd_celular_cliente)

	_Excel_Close($oExcel) ; fechar planilha

	;--------------------------------------CADASTRO-----------------------------------------
	ControlClick($tela_cadastro_clientes,"","TMyEdit7", "left",1,54, 13); Clica na Aba Manutenção
	Sleep(1000)


	ControlSend($tela_cadastro_clientes, "", "TMyEdit7", $sResultNome_cliente)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TJvDateEdit2", $sResultData_nascimento)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TMyEdit5", $sResultCpf_cliente)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TMyEdit4", $sResultIdentidade_cliente)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TMyEdit3", $sResultCartao_identidade_cliente)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TMyEdit2", $sResultContato_representante_cliente)
	Send("{enter}")
	Sleep(1000)
	;------------------------------------CEP------------------------------------------

	ControlClick($tela_cadastro_clientes,"","TMyEdit36", "left", 1)
	ControlSend($tela_cadastro_clientes,"","TMyEdit36",$cep_residencial_cliente)
	Send("{enter}")
	Sleep(1000)




	;------------------------------------ENDEREÇO RESIDENCIAL-------------------------

	ControlClick($tela_cadastro_clientes, "", "TMyEdit39", "left", 1)
	ControlSend($tela_cadastro_clientes, "", "TMyEdit39", $sResultEndereco_residencial_cliente)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TMyEdit24", $sResultNumero_residencial_cliente)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TMyEdit22", $sResultComplemento_residencial_cliente)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TMyEdit38", $sResultBairro_residencial_cliente)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TMyEdit18", $sResultReferencia_residencial_cliente)
	Send("{enter}")
	Sleep(1000)


	;-----------------------------------OUTROS CONTATOS---------------------------------------
	ControlSend($tela_cadastro_clientes, "", "TMyEdit28", $sResultTelefone_fixo_cliente)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TMyEdit16", $sResultDdd_celular_cliente )
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TMyEdit27", $sResultCelular_cliente )
	Send("{enter}")
	Sleep(1000)

	ControlSend($tela_cadastro_clientes, "", "TMyEdit19", $sResultDdd_telefone2_cliente)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TMyEdit26", $sResultTelefone2_cliente)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TMyEdit25", $sResultEmail_cliente)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TMyEdit13", $sResultContato_rede_social_cliente)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TMyEdit14", $sResultFacebook_cliente)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TMyEdit15", $sResultHomepage_cliente)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TMyEdit16", $sResultCaixa_postal_cliente)
	Send("{enter}")
	Sleep(1000)


	;-------------------------------COMPLEMENTO------------------------------------------------------
	ControlClick($tela_cadastro_clientes,"","TPageControl1", "left",1,139,13); Clica na Aba Complemento
	Sleep(1000)

	ControlSend($tela_cadastro_clientes, "", "TMyEdit10", $sResultNome_pai_cliente) ; Muda o que tiver no campo pra o que eu definir (Linha 17)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TMyEdit11",$sResultNome_mae_cliente)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TMyEdit18", $sResultReferencia1_cliente)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TMyEdit17", $sResultReferencia2_cliente)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TMyEdit16", $sResultReferencia3_cliente)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TMyEdit9", $sResultInscricao_municipal_cliente)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TMyEdit15", $sResultCampo_adicional1_cliente)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TMyEdit14", $sResultCampo_adicional2_cliente)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TMyEdit13", $sResultCampo_adicional3_cliente)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TMyEdit12", $sResultCampo_adicional4_cliente)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TComboBox1", $sResultSexo_cliente)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TMemo1", $sResultObservacoes_cliente)
	Send("{enter}")
	Sleep(1000)

;--------------------------------------------------FINANCEIRO----------------------------------------------------
	ControlClick($tela_cadastro_clientes,"","TPageControl1", "left",1,232,16); Clica Aba Financeiro
	Sleep(1000)

	ControlSend($tela_cadastro_clientes, "", "TMyEdit14", $sResultLimite_credito_cliente)
	Send("{enter}")
	Sleep(1000)
	;ControlSend($tela_cadastro_clientes, "", "TMyEdit13",$sResultConvenio_numero)
	;Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TComboBox4", $sResultPreco_cliente)
	Send("{enter}")
	Sleep(1000)
	ControlClick($tela_cadastro_clientes,"","TMyEdit10", "left",1,48, 17) ;Numero Vendedor Padrão
	ControlSend($tela_cadastro_clientes, "", "TMyEdit10", $sResultVendedor_padrao_cliente)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TComboBox3", $sResultSituacao_cliente)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TComboBox2", $sResultTipo_cliente)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TCheckBox1", $sResultPrograma_fidelidade_cliente)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TCheckBox2", $sResultBloquear_inadimplente_cliente)
	Send("{enter}")
	Sleep(1000)
	ControlSend($tela_cadastro_clientes, "", "TComboBox1", $sResultSetor_cliente)
	Send("{enter}")
	Sleep(1000)

;------------------------------------------------SALVAR-------------------------------------------
	ControlClick($tela_cadastro_clientes,"","TJvXPButton2","left",1) ;Salvar Cadastro
	Send("{enter}")
	Sleep(1000)

	MsgBox(0, "teste", "clique em ok para editar")
	ControlClick($tela_cadastro_clientes,"","TJvXPButton3","left",1) ;Editar
	Send("{enter}")


	MsgBox(0, "teste", "clique em ok para Excluir")
	ControlClick($tela_cadastro_clientes,"","TJvXPButton10","left",1) ;Excluir
	Send("{enter}")
	Sleep(1000)

EndFunc