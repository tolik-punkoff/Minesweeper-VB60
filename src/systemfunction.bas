Attribute VB_Name = "systemfunction"
Option Explicit 'контроль объ€влени€ переменных

Public Function FileExist(FileEx$) As Boolean
' проверка наличи€ файла
Dim ff As Long ' переменна€ под свободный файловый номер
On Error GoTo 10 ' если ошибка - то файла нет, либо зан€т
ff = FreeFile() ' находим свободный файловый номер
Open FileEx$ For Input As ff ' пытаемс€ открыть файл дл€ чтени€
Close ff ' закрываем
FileExist = True ' если ошибки до сюда не случилось, файл есть - уст. зн ф-ии в TRUE
Exit Function ' выходим из функции
10 FileExist = False ' иначе устанавливаем значение в FALSE
End Function

