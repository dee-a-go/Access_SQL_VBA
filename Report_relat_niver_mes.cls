VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_relat_niver_mes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Report_Load()
Select Case nUMERO_MES
        Case 1:
            nome_mes = "Janeiro"
        Case 2:
            nome_mes = "Fevereiro"
        Case 3:
            nome_mes = "Março"
        Case 4:
            nome_mes = "Abril"
        Case 5:
            nome_mes = "Maio"
        Case 6:
            nome_mes = "Junho"
        Case 7:
            nome_mes = "Julho"
        Case 8:
            nome_mes = "Agosto"
        Case 9:
            nome_mes = "Setembro"
        Case 10:
            nome_mes = "Outubro"
        Case 11:
            nome_mes = "Novembro"
        Case 12:
            nome_mes = "Dezembro"
            
        Case Is = "inválido"
            
End Select
MES_REF = UCase(nome_mes)

End Sub
