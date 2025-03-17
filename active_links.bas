Sub MakeLinksActiveInSelection()  
    Dim rng As Range  
    Dim cell As Range  
    
    ' Получаем выделенный диапазон  
    Set rng = Selection  
    
    ' Проверяем, что выделение не пустое  
    If rng Is Nothing Then  
        MsgBox "Пожалуйста, выделите диапазон ячеек с текстом ссылок."  
        Exit Sub  
    End If  
    
    ' Проходим по каждому элементу выделенного диапазона  
    For Each cell In rng  
        ' Проверяем, что ячейка не пустая  
        If Not IsEmpty(cell.Value) Then  
            ' Добавляем гиперссылку, предполагая, что текст является URL  
            On Error Resume Next ' Игнорируем возможные ошибки (например, если ссылка уже есть)  
            cell.Hyperlinks.Add Anchor:=cell, _  
                               Address:=cell.Value, _  
                               TextToDisplay:=cell.Value  
            On Error GoTo 0 ' Возвращаем обработку ошибок к стандартной  
        End If  
    Next cell  
    
End Sub
