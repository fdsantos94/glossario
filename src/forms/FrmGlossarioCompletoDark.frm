VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmGlossarioCompletoDark 
   Caption         =   "MIS | Management Information System"
   ClientHeight    =   9435.001
   ClientLeft      =   156
   ClientTop       =   588
   ClientWidth     =   12024
   OleObjectBlob   =   "FrmGlossarioCompletoDark.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmGlossarioCompletoDark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
    Dim valor As String
    Dim ItemSelecionado As Integer
    
    ' Texto da coluna da linha selecionada
    valor = Me.ListBox1.Value
    
    ' Nº indice da coluna da linha selecionada
    ItemSelecionado = Me.ListBox1.ListIndex
    
    ' Definição de valores para o segundo FrmGlossarioIndicador
    With FrmGlossarioDetalheDark
        .TxtIndicador.Value = Me.ListBox1.List(ItemSelecionado, 0)
        .TxtConceito.Value = Me.ListBox1.List(ItemSelecionado, 2)
        .TxtCalculo.Value = Me.ListBox1.List(ItemSelecionado, 1)
    End With
    
    FrmGlossarioDetalheDark.Show

End Sub

Private Sub TxtPesquisa_Change()
    
    ' condicional para verificar se o usuario está na lista de problemas
    ' caso esteja, retornará true e aplicará as configurações
    ' caso contrário, manterá a configuração original
    If GlossarioConfiguracoes.fn_ican = True Then
        ' Validação para efeito do texto exemplo
        If TxtPesquisa.Text = Empty Then
        
            LbPesquisa.Caption = "     Procure um indicador aqui..."
            TxtPesquisa.ForeColor = RGB(255, 255, 255)
            LbPesquisa.ForeColor = RGB(255, 255, 255)
        
        Else
        
            LbPesquisa.Caption = Empty
        
        End If
        
    Else
    
        ' Validação para efeito do texto exemplo
        If TxtPesquisa.Text = Empty Then
        
            LbPesquisa.Caption = "       Procure um indicador aqui..."
            TxtPesquisa.ForeColor = RGB(255, 255, 255)
            LbPesquisa.ForeColor = RGB(255, 255, 255)
        
        Else
        
            LbPesquisa.Caption = Empty
        
        End If
        
    End If
    
    ' Local onde é direcionado o texto pesquisado
    Glossario.Range("F2").Value = TxtPesquisa.Value
    
    Call GlossarioConfiguracoes.Filtro_Avancado_Indicador

End Sub

Private Sub UserForm_Initialize()
    
    Dim Linha           As Integer
    Dim Base            As Range
    Dim NomePlanilha    As String
    Dim Coluna1         As Integer
    Dim Coluna2         As Integer
    Dim Coluna3         As Integer
    Dim Usuario         As Boolean
    
    ' condicional para verificar se o usuario está na lista de problemas
    ' caso esteja, retornará true e aplicará as configurações
    ' caso contrário, manterá a configuração original
    If GlossarioConfiguracoes.fn_ican = True Then
    
        ' Define dimensões do form principal
        With Me
            .Top = 0
            .Height = 408
            .Width = 397.8
        End With
        
        ' Definição dos botões do form
        With Me
            .TxtPesquisa.Height = 25.5
            .TxtPesquisa.Left = 6
            .TxtPesquisa.Top = 36
            .TxtPesquisa.Width = 168
            .TxtPesquisa.BackStyle = fmBackStyleTransparent
            .TxtPesquisa.ForeColor = RGB(255, 255, 255)
            
            .LbPesquisa.Height = 29.25
            .LbPesquisa.Left = 0
            .LbPesquisa.Top = 39 '36
            .LbPesquisa.Width = 186
            .LbPesquisa.BackStyle = fmBackStyleTransparent
            .LbPesquisa.Caption = "     Procure um indicador aqui..."
            .LbPesquisa.Font.Size = 11
            .LbPesquisa.ForeColor = RGB(255, 255, 255)
            
            .ListBox1.Height = 316.3
            .ListBox1.Left = 12
            .ListBox1.Top = 66
            .ListBox1.Width = 363
        End With
    
    Else
    
        ' Define configurações das caixas de busca
        With Me
            .LbPesquisa.BackStyle = fmBackStyleTransparent
            .LbPesquisa.Caption = "       Procure um indicador aqui..."
            .TxtPesquisa.BackStyle = fmBackStyleTransparent
            .LbPesquisa.ForeColor = RGB(255, 255, 255)
            .TxtPesquisa.ForeColor = RGB(255, 255, 255)
        End With
    
    End If
        
    ' Pega o nome da planilha que contem os dados do glossário
    NomePlanilha = "'" & Glossario.Name & "'!"
    
    ' Conta o total de linhas
    Linha = Glossario.Range("A1").CurrentRegion.Rows.Count
    
    ' Define o intervalo das bases
    Set Base = Glossario.Range(Glossario.Cells(2, 1), Glossario.Cells(Linha, 3))
    
    ' Auto ajusta as colunas
    Glossario.Columns("A:C").AutoFit
    
    ' Pega a largura de cada coluna + um acrescimo para melhor visibilidade no form
    Coluna1 = Glossario.Columns("A:A").Width + 10
    Coluna2 = Glossario.Columns("B:B").Width + 10
    Coluna3 = Glossario.Columns("C:C").Width + 10
    
    ' Configura a list box do FrmGlossario
    With Me
        .ListBox1.ColumnCount = 2
        .ListBox1.RowSource = NomePlanilha & Base.Address ' Pega o nome da planilha e o endereço da base preenchida
        .ListBox1.Font.Size = 10
        .ListBox1.Font.Name = "Arial Nova Light"
        .ListBox1.ColumnWidths = Coluna1 & "pt;" & Coluna2 & "pt;" & Coluna3 & "pt" '"30 pt;150 pt;150 pt;
        .ListBox1.ColumnHeads = True
        .ListBox1.ForeColor = RGB(24, 23, 23)
    End With
    
End Sub

