Sub CalculateContourArea()
    Dim s As Shape
    Dim dArea As Double

    ' Verifica se há uma seleção
    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Por favor, selecione um contorno primeiro."
        Exit Sub
    End If

    ' Obtém a forma selecionada
    Set s = ActiveSelection.Shapes(1)

    ' Verifica se a forma é um contorno
    If Not s.Type = cdrCurveShape Then
        MsgBox "A forma selecionada não é um contorno."
        Exit Sub
    End If

    ' Calcula a área do contorno
    dArea = s.Curve.Area

    ' Também exibe uma mensagem com a área
    MsgBox "Área do contorno: " & dArea & " unidades quadradas"
End Sub
