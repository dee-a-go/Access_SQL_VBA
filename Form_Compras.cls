VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Compras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
    Call conectar_banco
    
    op_fem.Enabled = False
    op_masc.Enabled = False
    op_fem.Value = False
    op_masc.Value = False
    op_p.Enabled = False
    op_m.Enabled = False
    op_g.Enabled = False
    cmb_modelo.Enabled = False
    cmb_peca.RowSource = ""
    cmb_modelo.RowSource = ""
    txt_cpf.SetFocus
    
    With cmb_peca
        .AddItem ("Camisas")
        .AddItem ("Calças")
        .AddItem ("Shorts")
    End With

End Sub

Private Sub cmb_peca_Change()
    cmb_modelo.RowSource = ""
    op_fem.Enabled = True
    op_masc.Enabled = True
    cmb_modelo.Value = ""
    
    If cmb_peca.Value = "Calças" Then
        sql = "select CALCAS from tb_calcas"
        op_fem.Enabled = False
        op_masc.Enabled = False
        op_fem.Value = False
        op_masc.Value = False
        cmb_modelo.Enabled = True
        tb_selecionada = "tb_calcas"
        tipo_peca = "CALCAS"
        
    ElseIf cmb_peca.Value = "Shorts" And op_masc.Value Then
        sql = "select SHORTS from tb_shortsM"
        cmb_modelo.Enabled = True
        tb_selecionada = "tb_shortsM"
        tipo_peca = "Shorts"
        
    ElseIf cmb_peca.Value = "Shorts" And op_fem.Value Then
        sql = "select SHORTS from tb_shortsF"
        cmb_modelo.Enabled = True
        tb_selecionada = "tb_shortsF"
        tipo_peca = "Shorts"
        
    ElseIf cmb_peca.Value = "Camisas" And op_masc.Value Then
        sql = "select CAMISAS from tb_camisaM"
        cmb_modelo.Enabled = True
        tb_selecionada = "tb_camisaM"
        tipo_peca = "Camisas"
        
    ElseIf cmb_peca.Value = "Camisas" And op_fem.Value Then
        sql = "select CAMISAS from tb_camisaF"
        cmb_modelo.Enabled = True
        tb_selecionada = "tb_camisaF"
        tipo_peca = "Camisas"
        
    Else
        cmb_modelo.RowSource = ""
    End If

    cmb_modelo.RowSource = sql
End Sub

Private Sub cmb_modelo_Change()
    op_p.Enabled = True
    op_m.Enabled = True
    op_g.Enabled = True
    peca = cmb_modelo.Value
    sql = "select Valor from " & tb_selecionada & " where " & tipo_peca & " = '" & peca & "'"
    Call validar_leitura
    preco = rs!Valor
    lbl_valor.Caption = "R$ " & preco
End Sub

Private Sub op_fem_Click()
    op_masc.Value = False
    Call cmb_peca_Change
End Sub

Private Sub op_masc_Click()
    op_fem.Value = False
    Call cmb_peca_Change
End Sub

Private Sub op_g_Click()
    op_m.Value = False
    op_p.Value = False
    tamanho = "G"
End Sub

Private Sub op_m_Click()
    op_g.Value = False
    op_p.Value = False
    tamanho = "M"
End Sub

Private Sub op_p_Click()
    op_m.Value = False
    op_g.Value = False
    tamanho = "P"
End Sub

Private Sub btn_comprar_Click()
    If IsNull(txt_cpf) Or IsNull(cmb_peca) Or IsNull(cmb_modelo) Or (IsNull(op_p) And IsNull(op_m) And IsNull(op_g)) Then
        MsgBox ("Preencha todos os campos"), vbInformation + vbOKOnly, "Atenção"
        Exit Sub
    End If
    
    sql = "select * from tb_clientes where cpf = '" & txt_cpf & "'"
    Call validar_leitura
    If rs.EOF = True Then
        MsgBox ("CPF Inexistente!"), vbInformation + vbOKOnly, "Atenção"
    Else
        sql = "insert into tb_vendas values ('" & txt_cpf & "', '" & peca & "', '" & preco & "', '" & tamanho & "', '" & Date & "')"
        db.Execute (UCase(sql))
        MsgBox ("Pedido realizado com sucesso!!"), vbInformation + vbOKOnly, "AVISO"
        txt_cpf = ""
        cmb_peca.Value = ""
        cmb_modelo.Value = ""
        op_p.Value = False
        op_m.Value = False
        op_g.Value = False
        op_p.Enabled = False
        op_m.Enabled = False
        op_g.Enabled = False
        op_masc.Enabled = False
        op_fem.Enabled = False
        op_masc.Value = False
        op_fem.Value = False
        cmb_modelo.Enabled = False
        cmb_modelo.RowSource = ""
        txt_cpf.SetFocus
        lbl_valor.Caption = ""
    End If
End Sub
