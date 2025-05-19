Attribute VB_Name = "modIconeParaLimparTabelas"
Option Explicit

' Exemplo de uso:
' - Basta apenas deixar a planilha com as tabelas ativa e executar a sub 'GerarBotõesLimpar'

Public Sub GerarBotõesLimpar()
    
    'Declaração de variavies
    Dim Tabela          As ListObject
    Dim NomeTabela      As String
    Dim LocalIcone      As Boolean
    Dim CelulaFormula   As Range
    Dim Icone           As Object
    Dim IconeShp        As Shape
    Dim Resposta        As VbMsgBoxResult
    Dim RespostaFormula As VbMsgBoxResult
    
    'Verifica se existe shapes antes de limpar os botões
    If ActiveSheet.Shapes.Count > 0 Then
        'Deletando os botões existentes
        For Each IconeShp In ActiveSheet.Shapes
            If Right(IconeShp.Name, 4) = "btlp" Then
                IconeShp.Delete
            End If
        Next
    End If
    'Desativa o modo de copiar para evitar problemas
    Application.CutCopyMode = False
    
    'Verifica a quantidade de tabela na planilha ativa
    If ActiveSheet.ListObjects.Count = 0 Then
        MsgBox "Não foi possivel encontrar nenhuma planilha para adicionar o icone!", vbOKOnly + vbCritical, "Atenção"
        Exit Sub
    End If
    
    'Iterando por todas as tabelas da planilha ativa
    For Each Tabela In ActiveSheet.ListObjects

        NomeTabela = Tabela.Name '> Salvando o nome da tabela
        
        'Verifica se a tabela está na coluna 1
        If Tabela.Range.Cells(1, 1).Column = 1 Then
            'Se estiver então pergunta para o usurio se quer adicionar uma nova coluna _
            para um melhor ajuste do icone de borracha
            Resposta = MsgBox("Notamos que a tabela '" & NomeTabela & "' está na coluna 1, deseja " & _
                             "adicionar uma coluna para gerar o botão de limpar?", vbCritical + vbYesNo, "Atenção")
            
            'Verifica a resposta do usuario
            If Resposta = vbNo Or Resposta = vbAbort Then
                LocalIcone = True '> indica que o icone irá aparecer emcima da tabela
            ElseIf Resposta = vbYes Then
                Tabela.Range.Columns(1).EntireColumn.Insert '> Adiciona a nova coluna
            End If
            
        End If
        
        'Adiciona o icone
        On Error Resume Next '> Tolera os error adiante
        Set Icone = ActiveSheet.Pictures.Insert("https://cdn.hubblecontent.osi.office.net/icons/publish/icons_eraser/eraser.svg")
        On Error GoTo 0 '> Retorna depois de um error
        
        'Verifica se foi ou não possivel adicionar o icone
        If Icone Is Nothing Then
            MsgBox "Erro ao carregar o ícone. Verifique sua conexão com a internet.", vbCritical, "Erro"
            Exit Sub
        End If
        
        'Seleção para trabalhar apenas com o objeto icone
        With Icone
            .Name = NomeTabela & "btlp" '> Define o nome do icone
            .Height = Tabela.Range.Cells(1, 1).Height '> Define a altura do icone
            
            'Verifica o local do icone com base na resposta do usuario
            If LocalIcone = True Then
                .Left = Tabela.Range.Cells(1, 1).Left
                .Top = Tabela.Range.Cells(1, 1).Top
                LocalIcone = False
            Else
                .Left = Tabela.Range.Cells(1, 1).Left - Icone.Height
                .Top = Tabela.Range.Cells(1, 1).Top
            End If
        End With
       
        'Setando o iconeshape para usar o metodo onAction
        Set IconeShp = ActiveSheet.Shapes(Icone.Name)
        
        'Verifica se existe formulas na tabela
        On Error Resume Next '> Tolera os erros
        Set CelulaFormula = Tabela.DataBodyRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0 '> Retora depois de um erro
        
        'Se encontrar algum erro pergunta para o usuario se ele deseja ter o sub que deleta apenas células sem formulas
        If Not CelulaFormula Is Nothing Then
            RespostaFormula = MsgBox("Notamos que a tabela '" & NomeTabela & "' contem formulas,  " & _
                             "deseja limpar apenas os as celulas sem formulas?" & vbNewLine & vbNewLine & _
                             "Ao clicar em 'Sim' irá apenas limpar as celulas sem formulas." & vbNewLine & _
                             "Ao clicar em 'Não' irá limpar todas as celulas incluindo as celulas com formulas.", vbExclamation + vbYesNo, "Atenção")
            Set CelulaFormula = Nothing
        End If
        
        'Verifica a resposta do usuario para declara a sub correta para cada tabela
        If RespostaFormula = vbYes Then
            IconeShp.OnAction = "LimparTabelaComFormulas" '> Atribui a macro LimparTabelaComFormulas para o icone
        Else
            IconeShp.OnAction = "LimparTabela" '> Atribui a macro LimparTabela para o icone
        End If
    Next Tabela

End Sub

'Limpa a tabela
Sub LimparTabela()

    'Declaração de variaveis
    Dim Tabela          As ListObject
    
    '> On error para quando não encontrar a tabela
    On Error GoTo TabelaNãoEncontrada
    
    'Seta a tebela com base no nome do icone
    Set Tabela = ActiveSheet.ListObjects(Left(Application.Caller, Len(Application.Caller) - 4))
    Tabela.DataBodyRange.ClearContents '> Limpa o conteudo da tabela
    Tabela.DataBodyRange.Cells(1, 1).Select '> Seleciona a primeira célula da tabela
    Exit Sub '> Encerra a sub
    
'Tratamento de erros
TabelaNãoEncontrada:
    'Declaração de variaveis
    Dim Resposta    As VbMsgBoxResult
    
    'Mostra o problema e as soluções para o usuario
    Resposta = MsgBox("Não foi possivel encontrar a tabela " & Left(Application.Caller, Len(Application.Caller) - 4) & "!" & vbNewLine & vbNewLine & _
            "Soluções:" & vbNewLine & vbNewLine & "   1 - Renomear a planilha ao lado para '" & Left(Application.Caller, Len(Application.Caller) - 4) & "'. " & vbNewLine & _
            "   2 - Gerar os botões novamente." & vbNewLine & vbNewLine & "Deseja gerar os botões novamente? ", vbCritical + vbYesNo, "Atenção")
    'Gera novos botões se o usuario concordar
    If Resposta = vbYes Then
        Call GerarBotõesLimpar '> Chama a sub para gerar os botões
        Exit Sub '> Encerra a sub
    End If
End Sub

'Limpa as colunas que não contem formulas
Sub LimparTabelaComFormulas()
    'Declaração de variáveis
    Dim Tabela          As ListObject
    Dim Coluna          As ListColumn
    Dim ColunaFormula   As Range

    
    '> On error para quando não encontrar a tabela
    On Error GoTo TabelaNãoEncontrada
    
    'Seta a tabela com base no nome do ícone
    Set Tabela = ActiveSheet.ListObjects(Left(Application.Caller, Len(Application.Caller) - 4))
    
    'Itera sobre cada coluna da tabela
    For Each Coluna In Tabela.ListColumns
        Set ColunaFormula = Nothing '> Redefine as variavel
        
        'Tentando obter células com fórmulas na coluna
        On Error Resume Next
        Set ColunaFormula = Coluna.DataBodyRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
        
        'Se não encontrar células com fórmulas, limpa a coluna
        If ColunaFormula Is Nothing Then
            Coluna.DataBodyRange.ClearContents
        End If
        
    Next Coluna
    'Seleciona a primeira célula da tabela
    Tabela.DataBodyRange.Cells(1, 1).Select
    Exit Sub '> Encerra a sub
    
'Tratamento de erros
TabelaNãoEncontrada:
    'Declaração de variáveis
    Dim Resposta    As VbMsgBoxResult
    
    'Mostra o problema e as soluções para o usuário
    Resposta = MsgBox("Não foi possível encontrar a tabela " & Left(Application.Caller, Len(Application.Caller) - 4) & "!" & vbNewLine & vbNewLine & _
            "Soluções:" & vbNewLine & vbNewLine & "   1 - Renomear a planilha ao lado para '" & Left(Application.Caller, Len(Application.Caller) - 4) & "'. " & vbNewLine & _
            "   2 - Gerar os botões novamente." & vbNewLine & vbNewLine & "Deseja gerar os botões novamente? ", vbCritical + vbYesNo, "Atenção")
    
    'Gera novos botões se o usuário concordar
    If Resposta = vbYes Then
        Call GerarBotõesLimpar '> Chama a sub para gerar os botões
        Exit Sub '> Encerra a sub
    End If
End Sub

