VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
    Call conectar_banco
End Sub

Private Sub btn_encerrar_Click()
    resp = MsgBox("Deseja realmente encerrar o sistema", vbExclamation + vbYesNo, "Atenção")
    If resp = vbYes Then
        DoCmd.Quit acQuitSaveAll
    End If
End Sub

Private Sub btn_excluir_Click()
    If IsNull(txt_cpf) Or _
       IsNull(txt_nome) Or _
       IsNull(txt_nascimento) Or _
       IsNull(txt_bairro) Or _
       IsNull(txt_cidade) Or _
       IsNull(txt_uf) Or _
       IsNull(txt_numero_casa) Or _
       IsNull(txt_celular) Or _
       IsNull(txt_email) Or _
       IsNull(txt_cep) Or _
       IsNull(txt_endereco) Then
        MsgBox ("Preencha todos os campos"), vbInformation + vbOKOnly, "Atenção"
        Exit Sub
    Else
        resp = MsgBox("Deseja realmente excluir o registro?", vbQuestion + vbYesNo, "ATENÇÃO")
        If resp = vbYes Then
            sql = "delete * from tb_clientes where cpf='" & txt_cpf & "'"
            db.Execute (sql)
            MsgBox ("Registro excluído com sucesso!"), vbInformation + vbOKOnly, "AVISO"
            Call limpar_cadastro
            txt_cpf.SetFocus
        End If
    End If
End Sub

Private Sub btn_gravar_Click()

    If IsNull(txt_cpf) Or _
       IsNull(txt_nome) Or _
       IsNull(txt_nascimento) Or _
       IsNull(txt_bairro) Or _
       IsNull(txt_cidade) Or _
       IsNull(txt_uf) Or _
       IsNull(txt_numero_casa) Or _
       IsNull(txt_celular) Or _
       IsNull(txt_email) Or _
       IsNull(txt_cep) Or _
       IsNull(txt_endereco) Then
        MsgBox ("Preencha todos os campos"), vbInformation + vbOKOnly, "Atenção"
        Exit Sub
    Else
        sql = "select * from tb_clientes where cpf='" & txt_cpf & "'"
        Call validar_leitura
        If rs.EOF = True Then
            sql = "insert into tb_clientes values ('" & txt_cpf & "','" & txt_nome & "','" & txt_nascimento & "','" & txt_cep & "','" & txt_endereco & "','" & txt_numero_casa & "', '" & txt_bairro & "','" & txt_cidade & "','" & txt_uf & "','" & txt_celular & "','" & txt_email & "')"
            db.Execute (UCase(sql))
            MsgBox ("Dados gravados com sucesso"), vbInformation + vbOKOnly, "parabéns"
        Else
            sql = "update tb_clientes set nome = '" & txt_nome & "', " & _
                                         "data_nascimento = '" & txt_nascimento & "', " & _
                                         "cep = '" & txt_cep & "', endereco = '" & txt_endereco & "', " & _
                                         "num_comp = '" & txt_numero_casa & "', " & _
                                         "bairro = '" & txt_bairro & "', cidade = '" & txt_cidade & "', " & _
                                         "uf = '" & txt_uf & "', celular = '" & txt_celular & "', " & _
                                         "email = '" & txt_email & "' where cpf = '" & txt_cpf & "'"
            db.Execute (UCase(sql))
            MsgBox ("Dados alterados com sucesso!"), vbInformation + vbOKOnly, "AVISO"
            
        End If
    End If
    
      Call limpar_cadastro
End Sub

Private Sub btn_relatorios_Click()
    DoCmd.Close
    DoCmd.OpenForm ("Relatorios"), acNormal
End Sub

Private Sub txt_cep_LostFocus()
    sql = "select * from tb_cep where cep ='" & txt_cep & "'"
    Call validar_leitura
    If rs.EOF = False Then
        txt_endereco = rs.Fields(1).Value
        txt_cidade = rs.Fields(2).Value
        txt_bairro = rs.Fields(3).Value
        txt_uf = rs.Fields(4).Value
        txt_numero_casa.SetFocus
    Else
        MsgBox ("CEP errado!"), vbInformation + vbOKOnly, "Atenção"
    End If
End Sub


Private Sub txt_cpf_LostFocus()
    If IsNull(txt_cpf) Then
       Exit Sub
    Else
        sql = "select * from tb_clientes where cpf='" & txt_cpf & "'"
        Call validar_leitura
        If rs.EOF = False Then
            txt_nome = rs.Fields(1).Value
            txt_nascimento = rs.Fields(2).Value
            txt_cep = rs.Fields(3).Value
            txt_endereco = rs.Fields(4).Value
            txt_numero_casa = rs.Fields(5).Value
            txt_bairro = rs.Fields(6).Value
            txt_cidade = rs.Fields(7).Value
            txt_uf = rs.Fields(8).Value
            txt_celular = rs.Fields(9).Value
            txt_email = rs.Fields(10).Value
            BTN_EXCLUIR.Enabled = True
        Else
            MsgBox ("CPF: " & txt_cpf & " não cadastrado. Por favor, efetue o cadastro."), vbInformation + vbOKOnly, "AVISO"
            txt_nome.SetFocus
        End If
    End If
End Sub
