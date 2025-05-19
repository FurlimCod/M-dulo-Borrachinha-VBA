Attribute VB_Name = "modIconeParaLimparTabelas"
Option Explicit

' Exemplo de uso:
' - Basta apenas deixar a planilha com as tabelas ativa e executar a sub 'GerarBot�esLimpar'

Public Sub GerarBot�esLimpar()
    
    'Declara��o de variavies
    Dim Tabela          As ListObject
    Dim NomeTabela      As String
    Dim LocalIcone      As Boolean
    Dim CelulaFormula   As Range
    Dim Icone           As Object
    Dim IconeShp        As Shape
    Dim Resposta        As VbMsgBoxResult
    Dim RespostaFormula As VbMsgBoxResult
    
    'Verifica se existe shapes antes de limpar os bot�es
    If ActiveSheet.Shapes.Count > 0 Then
        'Deletando os bot�es existentes
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
        MsgBox "N�o foi possivel encontrar nenhuma planilha para adicionar o icone!", vbOKOnly + vbCritical, "Aten��o"
        Exit Sub
    End If
    
    'Iterando por todas as tabelas da planilha ativa
    For Each Tabela In ActiveSheet.ListObjects

        NomeTabela = Tabela.Name '> Salvando o nome da tabela
        
        'Verifica se a tabela est� na coluna 1
        If Tabela.Range.Cells(1, 1).Column = 1 Then
            'Se estiver ent�o pergunta para o usurio se quer adicionar uma nova coluna _
            para um melhor ajuste do icone de borracha
            Resposta = MsgBox("Notamos que a tabela '" & NomeTabela & "' est� na coluna 1, deseja " & _
                             "adicionar uma coluna para gerar o bot�o de limpar?", vbCritical + vbYesNo, "Aten��o")
            
            'Verifica a resposta do usuario
            If Resposta = vbNo Or Resposta = vbAbort Then
                LocalIcone = True '> indica que o icone ir� aparecer emcima da tabela
            ElseIf Resposta = vbYes Then
                Tabela.Range.Columns(1).EntireColumn.Insert '> Adiciona a nova coluna
            End If
            
        End If
        
        'Adiciona o icone
        On Error Resume Next '> Tolera os error adiante
        Set Icone = ActiveSheet.Pictures.Insert("https://cdn.hubblecontent.osi.office.net/icons/publish/icons_eraser/eraser.svg")
        On Error GoTo 0 '> Retorna depois de um error
        
        'Verifica se foi ou n�o possivel adicionar o icone
        If Icone Is Nothing Then
            MsgBox "Erro ao carregar o �cone. Verifique sua conex�o com a internet.", vbCritical, "Erro"
            Exit Sub
        End If
        
        'Sele��o para trabalhar apenas com o objeto icone
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
        
        'Se encontrar algum erro pergunta para o usuario se ele deseja ter o sub que deleta apenas c�lulas sem formulas
        If Not CelulaFormula Is Nothing Then
            RespostaFormula = MsgBox("Notamos que a tabela '" & NomeTabela & "' contem formulas,  " & _
                             "deseja limpar apenas os as celulas sem formulas?" & vbNewLine & vbNewLine & _
                             "Ao clicar em 'Sim' ir� apenas limpar as celulas sem formulas." & vbNewLine & _
                             "Ao clicar em 'N�o' ir� limpar todas as celulas incluindo as celulas com formulas.", vbExclamation + vbYesNo, "Aten��o")
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

    'Declara��o de variaveis
    Dim Tabela          As ListObject
    
    '> On error para quando n�o encontrar a tabela
    On Error GoTo TabelaN�oEncontrada
    
    'Seta a tebela com base no nome do icone
    Set Tabela = ActiveSheet.ListObjects(Left(Application.Caller, Len(Application.Caller) - 4))
    Tabela.DataBodyRange.ClearContents '> Limpa o conteudo da tabela
    Tabela.DataBodyRange.Cells(1, 1).Select '> Seleciona a primeira c�lula da tabela
    Exit Sub '> Encerra a sub
    
'Tratamento de erros
TabelaN�oEncontrada:
    'Declara��o de variaveis
    Dim Resposta    As VbMsgBoxResult
    
    'Mostra o problema e as solu��es para o usuario
    Resposta = MsgBox("N�o foi possivel encontrar a tabela " & Left(Application.Caller, Len(Application.Caller) - 4) & "!" & vbNewLine & vbNewLine & _
            "Solu��es:" & vbNewLine & vbNewLine & "   1 - Renomear a planilha ao lado para '" & Left(Application.Caller, Len(Application.Caller) - 4) & "'. " & vbNewLine & _
            "   2 - Gerar os bot�es novamente." & vbNewLine & vbNewLine & "Deseja gerar os bot�es novamente? ", vbCritical + vbYesNo, "Aten��o")
    'Gera novos bot�es se o usuario concordar
    If Resposta = vbYes Then
        Call GerarBot�esLimpar '> Chama a sub para gerar os bot�es
        Exit Sub '> Encerra a sub
    End If
End Sub

'Limpa as colunas que n�o contem formulas
Sub LimparTabelaComFormulas()
    'Declara��o de vari�veis
    Dim Tabela          As ListObject
    Dim Coluna          As ListColumn
    Dim ColunaFormula   As Range

    
    '> On error para quando n�o encontrar a tabela
    On Error GoTo TabelaN�oEncontrada
    
    'Seta a tabela com base no nome do �cone
    Set Tabela = ActiveSheet.ListObjects(Left(Application.Caller, Len(Application.Caller) - 4))
    
    'Itera sobre cada coluna da tabela
    For Each Coluna In Tabela.ListColumns
        Set ColunaFormula = Nothing '> Redefine as variavel
        
        'Tentando obter c�lulas com f�rmulas na coluna
        On Error Resume Next
        Set ColunaFormula = Coluna.DataBodyRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
        
        'Se n�o encontrar c�lulas com f�rmulas, limpa a coluna
        If ColunaFormula Is Nothing Then
            Coluna.DataBodyRange.ClearContents
        End If
        
    Next Coluna
    'Seleciona a primeira c�lula da tabela
    Tabela.DataBodyRange.Cells(1, 1).Select
    Exit Sub '> Encerra a sub
    
'Tratamento de erros
TabelaN�oEncontrada:
    'Declara��o de vari�veis
    Dim Resposta    As VbMsgBoxResult
    
    'Mostra o problema e as solu��es para o usu�rio
    Resposta = MsgBox("N�o foi poss�vel encontrar a tabela " & Left(Application.Caller, Len(Application.Caller) - 4) & "!" & vbNewLine & vbNewLine & _
            "Solu��es:" & vbNewLine & vbNewLine & "   1 - Renomear a planilha ao lado para '" & Left(Application.Caller, Len(Application.Caller) - 4) & "'. " & vbNewLine & _
            "   2 - Gerar os bot�es novamente." & vbNewLine & vbNewLine & "Deseja gerar os bot�es novamente? ", vbCritical + vbYesNo, "Aten��o")
    
    'Gera novos bot�es se o usu�rio concordar
    If Resposta = vbYes Then
        Call GerarBot�esLimpar '> Chama a sub para gerar os bot�es
        Exit Sub '> Encerra a sub
    End If
End Sub

