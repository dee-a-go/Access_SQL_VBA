VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
    If IsNull(logado) Then
        logado = False
    End If
End Sub

Private Sub btn_login_Click()
    Call conectar_banco

    If IsNull(txt_usuario) Or IsNull(txt_senha) Then
       MsgBox ("Preencha todos os campos"), vbExclamation + vbOKOnly, "Atenção"
       txt_usuario.SetFocus
    Else
        sql = "select * from tb_usuarios where (nome='" & txt_usuario & "' or email='" & txt_usuario & "') and senha='" & txt_senha & "'"
        Call validar_leitura
        If rs.EOF = False Then
           logado = True
           DoCmd.Close
           DoCmd.OpenForm ("cadastro"), acNormal
           DoCmd.SelectObject acTable, "", True
        Else
           MsgBox ("Conta Inválida!"), vbExclamation + vbOKOnly, "Atenção"
           txt_usuario = Empty
           txt_senha = Empty
           txt_usuario.SetFocus
        End If
    End If
End Sub

Private Sub Form_Close()
    If Not logado Then
        DoCmd.Quit acQuitSaveAll
    End If
End Sub
