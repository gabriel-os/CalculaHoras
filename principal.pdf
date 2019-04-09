Public dias_corridos As Integer
Public INDICE_ABA_HORARIOS As Integer
Public INDICE_ABA_ATUAL As Integer

'Pega a ultima aba com os horários de funcionamento
Sub ultimaAba()
    INDICE_ABA_HORARIOS = ThisWorkbook.Sheets.Count
End Sub


'Conta quantas propostas há numa determinada planilha
Function contaPropostas(indice As Integer) As Integer

    contaPropostas = WorksheetFunction.CountIf(ThisWorkbook.Sheets(indice).Range("A:A"), ">""")

End Function


Sub calculaDias()
    
    Dim wkb As Workbook
    
    Set wkb = ThisWorkbook
    
    dias_corridos = DateDiff("D", wkb.Sheets(1).Cells(2, 2), wkb.Sheets(1).Cells(2, 4)) + 1

End Sub

Sub teste()
    
    Dim wkb As Workbook
    
    Set wkb = ThisWorkbook
    
    Call ultimaAba
    INDICE_ABA_ATUAL = 1
    Call calculaSegundos(wkb.Sheets(1).Cells(2, 2), wkb.Sheets(1).Cells(2, 3), wkb.Sheets(1).Cells(2, 4), wkb.Sheets(1).Cells(2, 5))
End Sub

Function calculaSegundos(data_inicio As Date, hora_inicio As Date, data_fim As Date, hora_fim As Date) As Double
    Dim wkb As Workbook
    Dim coluna_datas As Range
    Dim r As Object
    Dim coluna_inicio, coluna_fim As Integer
    Dim delta_dias As Integer
    Dim contador As Integer
    Dim hFim_aba_horario, hInicio_aba_horario As Date
    Dim tempSegundos As Long
    
    Set wkb = ThisWorkbook
    Set coluna_datas = wkb.Sheets(INDICE_ABA_HORARIOS).Range("A:A")
    
    delta_dias = DateDiff("D", data_inicio, data_fim)
    
    
    If Left(wkb.Sheets(INDICE_ABA_ATUAL).Name, 3) = "For" Then
        coluna_inicio = 2
        coluna_fim = 3
    Else
        coluna_inicio = 4
        coluna_fim = 5
    End If
          
    If delta_dias <= 0 Then
        'Condição caso esteja com a data errada
        calculaSegundos = -1
        Exit Function
    ElseIf delta_dias = 0 Then
        'Condição caso a proposta seja realizada no mesmo dia
        calculaSegundos = DateDiff("S", hora_inicio, hora_fim) / (60 * 60 * 24)
        Exit Function
    ElseIf delta_dias = 1 Then
        'Condição caso a proposta seja realizada no dia seguinte
        For Each r In coluna_datas
            contador = contador + 1
            
            If r.Value = data_inicio Then
                
                hFim_aba_horario = CDate(wkb.Sheets(INDICE_ABA_HORARIOS).Cells(contador, coluna_fim).Value)
                tempSegundos = DateDiff("S", hora_inicio, hFim_aba_horario)
                
            End If
            
        Next r
    End If
    
    For Each r In coluna_datas
        
    Next r
End Function


