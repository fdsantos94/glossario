VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmGlossarioDetalheDark 
   Caption         =   "Detalhe Indicador"
   ClientHeight    =   4200
   ClientLeft      =   156
   ClientTop       =   576
   ClientWidth     =   8160
   OleObjectBlob   =   "FrmGlossarioDetalheDark.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmGlossarioDetalheDark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    
    ' condicional para verificar se o usuario está na lista de problemas
    ' caso esteja, retornará true e aplicará as configurações
    ' caso contrário, manterá a configuração original
    If GlossarioConfiguracoes.fn_ican = True Then
    
        ' Define dimensões do form
        With Me
            .Top = 0.6
            .Height = 196.8
            .Width = 337.2
            .Left = 0.6
        End With
        
        ' Define dos botões do form
        With Me
            .TxtIndicador.Height = 36
            .TxtIndicador.Left = 12
            .TxtIndicador.Top = 6
            .TxtIndicador.Width = 324
            .TxtIndicador.BackStyle = fmBackStyleTransparent
            .TxtIndicador.ForeColor = RGB(255, 255, 255)
        
            .TxtConceito.Height = 58
            .TxtConceito.Left = 12
            .TxtConceito.Top = 41
            .TxtConceito.Width = 300
            .TxtConceito.Font.Size = 11
            .TxtConceito.BackStyle = fmBackStyleTransparent
            .TxtConceito.ForeColor = RGB(255, 255, 255)
        
            .TxtCalculo.Height = 48
            .TxtCalculo.Left = 12
            .TxtCalculo.Top = 118
            .TxtCalculo.Width = 300
            .TxtCalculo.Font.Size = 11
            .TxtCalculo.BackStyle = fmBackStyleTransparent
            .TxtCalculo.ForeColor = RGB(255, 255, 255)
        End With
    
    Else
    
        ' Define configurações das caixas de texto
        With Me
            .TxtIndicador.BackStyle = fmBackStyleTransparent
            .TxtIndicador.ForeColor = RGB(255, 255, 255)
            
            .TxtCalculo.BackStyle = fmBackStyleTransparent
            .TxtCalculo.ForeColor = RGB(255, 255, 255)
            
            .TxtConceito.BackStyle = fmBackStyleTransparent
            .TxtConceito.ForeColor = RGB(255, 255, 255)
        End With
    
    End If

End Sub

