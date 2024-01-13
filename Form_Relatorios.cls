VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Relatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btn_gerar_Click()
    
    
    If IsNull(cmb_relatorios.Value) Then
        MsgBox ("Selecione um relatório"), vbExclamation + vbOKOnly, "ATENÇÃO"
        Exit Sub
    ElseIf cmb_relatorios.Value = "Listagem geral de clientes" Then
        nome_relatorio = "relat_clientes"
    ElseIf cmb_relatorios.Value = "Listagem de endereços por cidades e bairros" Then
        nome_relatorio = "relat_enderecos"
    ElseIf cmb_relatorios.Value = "Listagem de aniversariantes do mês" Then
        nome_relatorio = "relat_niver_mes"
    ElseIf cmb_relatorios.Value = "Listagem entre periodos (ano de nascimento)" Then
        nome_relatorio = "relat_periodo"
    ElseIf cmb_relatorios.Value = "Historico de vendas" Then
        nome_relatorio = "relat_vendas"
    End If
    

    DoCmd.OpenReport (nome_relatorio), acViewReport
End Sub

Private Sub Form_Load()
Call conectar_banco
    cmb_relatorios.RowSource = ""
    With cmb_relatorios
        .AddItem ("Listagem geral de clientes")
        .AddItem ("Listagem de endereços por cidades e bairros")
        .AddItem ("Listagem de aniversariantes do mês")
        .AddItem ("Listagem entre periodos (ano de nascimento)")
        .AddItem ("Historico de vendas")
    End With
End Sub
