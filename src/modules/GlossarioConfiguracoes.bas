Attribute VB_Name = "GlossarioConfiguracoes"
Option Explicit

Function fn_ican() As Boolean
    
    Dim Linha As Integer
    Dim i As Integer
    
    ' verifica o total de linhas
    Linha = Glossario.Range("AD1").CurrentRegion.Rows.Count + 1
    
    ' loop para verificar se o usuario está na lista de problemas
    ' caso esteja, retornará true e aplicará as configurações
    ' caso contrário, manterá a configuração original
    For i = 1 To Linha
    
        If UCase(Glossario.Range("AD" & i).Value) = UCase(Environ$("username")) Then
        
            fn_ican = True
                    
        End If
    
    Next

End Function

Sub ExibirGlossario()

    FrmGlossarioCompletoDark.Show
    
End Sub

Sub Filtro_Avancado_Indicador()
    
    Dim Base            As Range
    Dim Criterio        As Range
    Dim FiltroRealizado As Range
    Dim NomePlanilha    As String
    Dim Linha           As Integer
    
    ' Limpa base onde será o resultado do filtro avançado
    Glossario.Range("J2:L1048576").Clear
    
    ' Ultima linha
    Linha = Glossario.Range("A1").CurrentRegion.Rows.Count
    
    ' Define intervalo principal do glossário
    Set Base = Glossario.Range("A1").CurrentRegion
    
    ' Range de critério ou busca, local onde será despejado as palavras escritas no form
    Set Criterio = Glossario.Range("F1:H2")
    
    ' Adiciona o curinga "*" antes e depois do texto para simular a busca por like
    Glossario.Range("F2").Value = "*" & FrmGlossarioCompletoDark.TxtPesquisa.Value & "*"

    ' Local onde será despejado a base de acordo com o critério escrito
    Set FiltroRealizado = Glossario.Range(Glossario.Cells(2, 10), Glossario.Cells(Linha, 12))
    
    ' Pega o nome da planilha
    NomePlanilha = "'" & Glossario.Name & "'!"
    
    ' Local onde sera aplicado o filtro
    Base.AdvancedFilter xlFilterCopy, Criterio, Glossario.Range("J1:L1")
    
    ' Pega o endereço do filtro para carregar no listbox do formulário
    FrmGlossarioCompletoDark.ListBox1.RowSource = NomePlanilha & FiltroRealizado.Address

End Sub


