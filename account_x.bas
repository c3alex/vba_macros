Function СЧЕТХ(inputString As String, keywordPattern As String) As Long
    Dim lines() As String
    Dim line As Variant
    Dim regex As Object
    Dim matches As Object
    Dim sumX As Long
    Dim match As Variant
    
    sumX = 0 ' Инициализация переменной для хранения суммы чисел с "x" в конце
    
    ' Создаем объект регулярного выражения
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Настройка регулярного выражения для поиска числа с "x" в конце
    With regex
        .Global = True ' Чтобы найти все совпадения в строке
        .IgnoreCase = True ' Игнорируем регистр
        .Pattern = "(\d+)x\b" ' Регулярное выражение для поиска чисел с "x" в конце
    End With
    
    ' Заменяем возможные символы новой строки (vbCrLf, vbLf, и прочие) на один стандартный символ новой строки (vbCrLf)
    inputString = Replace(inputString, vbLf, vbCrLf) ' Заменяем одиночные символы новой строки
    inputString = Replace(inputString, vbCr, vbCrLf) ' В случае если используются только vbCr
    
    ' Разделяем входную строку на отдельные строки по символу новой строки
    lines = Split(inputString, vbCrLf)
    
    ' Отладочный вывод: показываем все строки, на которые разделена входная строка
'    Debug.Print "Все строки:"
'    For Each line In lines
'        Debug.Print line
'    Next line
    
    ' Проходим по каждой строке
    For Each line In lines
        ' Отладочный вывод: показываем текущую строку
'        Debug.Print "Текущая строка: " & line
        
        ' Проверяем, содержит ли строка ключевое слово
        If InStr(line, keywordPattern) > 0 And InStr(line, "или") = 0 Then
            ' Отладочный вывод: показываем, что строка содержит ключевое слово
'            Debug.Print "Ключевое слово найдено в строке: " & line
            
            ' Удаляем лишние пробелы для точной проверки
            line = Trim(line)
            
            ' Ищем все совпадения с числом и "x" в текущей строке
            Set matches = regex.Execute(line)
            
            ' Отладочный вывод: показываем найденные совпадения
'            If matches.count > 0 Then
'                Debug.Print "Найдено " & matches.count & " совпадений в строке."
'            Else
'                Debug.Print "Совпадений не найдено."
'            End If
            
            ' Суммируем найденные числа, если они соответствуют ключевому слову
            For Each match In matches
                ' Отладочный вывод: показываем найденное совпадение
'                Debug.Print "Найдено число с 'x': " & match.value
                
                ' Добавляем найденное число к общей сумме
                sumX = sumX + CLng(match.SubMatches(0))
            Next match
        Else
            ' Отладочный вывод: строка не содержит ключевого слова
'            Debug.Print "Ключевое слово не найдено в строке."
        End If
    Next line
    
'    ' Отладочный вывод: итоговая сумма
'    Debug.Print "Итоговая сумма: " & sumX
    
    ' Возвращаем результат
    СЧЕТХ = sumX
End Function
