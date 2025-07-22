# Comandos CMD ou Terminal

Instalar bibliotecas:
pip install pandas PySide6 numpy openpyxl

VBA Excel:
Private Sub Worksheet_Change(ByVal Target As Range)
    ' --- CONFIGURAÇÃO GERAL (COMBINADA) ---
    ' Coluna para a Prioridade (A, B, C, etc.)
    Const COLUNA_PRIORIDADE As String = "A"
    ' Linha onde os dados começam (geralmente 2, após cabeçalho)
    Const PRIMEIRA_LINHA_DADOS As Integer = 2
    
    ' Coluna para o Status (A=1, B=2, C=3, D=4, E=5, etc.)
    Const COLUNA_STATUS As Integer = 5   ' <<< Verifique se esta é sua coluna de Status
    ' Coluna para a Data do Status
    Const COLUNA_DATA As Integer = 6     ' <<< Verifique se esta é sua coluna de Data
    ' ---------------------------------------

    ' Sai da macro se mais de uma célula for alterada ou se for apagada
    If Target.Cells.Count > 1 Or Target.Value = "" Then Exit Sub

    ' --- LÓGICA DE PRIORIZAÇÃO AUTOMÁTICA ---
    ' Verifica se a mudança foi na coluna de Prioridade
    If Target.Column = Me.Range(COLUNA_PRIORIDADE & "1").Column Then
        ' Se o valor digitado for 1, reordena tudo
        If Target.Value = 1 Then
            Application.EnableEvents = False ' Desliga eventos
            
            Dim linhaEditada As Long
            linhaEditada = Target.Row
            
            ' Move a linha para o topo
            If linhaEditada > PRIMEIRA_LINHA_DADOS Then
                Me.Rows(linhaEditada).Cut
                Me.Rows(PRIMEIRA_LINHA_DADOS).Insert Shift:=xlDown
            End If
            
            ' Renumera toda a coluna de Prioridade
            Dim ultimaLinha As Long
            ultimaLinha = Me.Cells(Me.Rows.Count, COLUNA_PRIORIDADE).End(xlUp).Row
            
            Dim i As Long
            For i = PRIMEIRA_LINHA_DADOS To ultimaLinha
                Me.Cells(i, Me.Range(COLUNA_PRIORIDADE & "1").Column).Value = i - (PRIMEIRA_LINHA_DADOS - 1)
            Next i
            
            Me.Cells(PRIMEIRA_LINHA_DADOS, Me.Range(COLUNA_PRIORIDADE & "1").Column).Offset(1, 0).Select
            
            Application.EnableEvents = True ' Liga eventos
        End If
    End If

    ' --- LÓGICA DE DATA AUTOMÁTICA ---
    ' Verifica se a mudança foi na coluna de Status
    If Target.Column = COLUNA_STATUS Then
        ' Se o status for Concluído ou Cancelado, adiciona a data/hora
        If LCase(Target.Value) = "concluído" Or LCase(Target.Value) = "cancelado" Then
            Application.EnableEvents = False ' Desliga eventos
            
            Me.Cells(Target.Row, COLUNA_DATA).Value = Now()
            
            Application.EnableEvents = True ' Liga eventos
        End If
    End If
End Sub
