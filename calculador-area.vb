Sub CalculateContourArea()
    Dim formato As Shape
    Dim area As Double

    ' Verifica se há uma seleção
    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Por favor, selecione um contorno primeiro."
        Exit Sub
    End If

    ' Obtém a forma selecionada
    Set formato = ActiveSelection.Shapes(1)

    ' Verifica se a forma é um contorno
    If Not formato.Type = cdrCurveShape Then
        MsgBox "A forma selecionada não é um contorno."
        Exit Sub
    End If

    ' Calcula a área do contorno
    area = Round(formato.Curve.Area, 5)

    ' Também exibe uma mensagem com a área
    MsgBox "Área do contorno: " & area & " unidades quadradas"
End Sub
