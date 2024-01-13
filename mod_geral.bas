Attribute VB_Name = "mod_geral"
Option Compare Database
Public db As Database
Public rs As Recordset
Public sql, resp, nome_mes, nome_relatorio, peca, tb_pecas, tipo_peca, tb_selecionada, tabela_selecionada, preco, tamanho As String
Public logado As Boolean

Function conectar_banco()
    Set db = CurrentDb
End Function

Function validar_leitura()
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)
End Function

Function limpar_cadastro()
    With Form_cadastro
        .txt_cpf = ""
        .txt_nome = ""
        .txt_nascimento = ""
        .txt_cep = ""
        .txt_bairro = ""
        .txt_cidade = ""
        .txt_endereco = ""
        .txt_email = ""
        .txt_celular = ""
        .txt_numero_casa = ""
        .txt_uf = ""
    End With
End Function
