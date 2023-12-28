Attribute VB_Name = "saper0"
Option Explicit 'контроль объявления переменных
' Константы уровня Чайник
Public Const ChainMines = 10 'мины
Public Const ChainX = 10 'клетки по горизонтали
Public Const ChainY = 10 'клетки по вертикали
' Константы уровня Бывалый
Public Const BivalMines = 40 'Аналогично выше
Public Const BivalX = 16
Public Const BivalY = 16
' Константы уровня Крутой
Public Const CoolMines = 99 'Аналогично выше
Public Const CoolX = 30
Public Const CoolY = 16
' Переменные уровня Другой
Public OthMines As Long 'Аналогично выше
Public OthX As Long
Public OthY As Long

Public Mines(750) As Long ' Массив с указаниями где мины и где какие цифры в соседних с минами клетках
Public Flags(750) As Boolean ' Массив с указаниями где флаги
Public Sosedi(8) As Long ' Массив с номерами соседних клеток (д. функции GetSosedi)
Public FlagCtr As Integer ' счетчик флагов
Public Sub GetSosedi(Tek As Long, x As Long, Maximum As Long) 'Нумерация с 1
' GetSosedi(Номер_клетки_в_массиве_Mines_для_которой_нужны_соседи, Кол-во_клеток_по_горизонтали, кол-во_клеток_на_поле)
Dim I As Byte
For I = 0 To 8 'Почистим массив с соседями, в него могли нагадить!
    Sosedi(I) = 0
Next I
'Проверить на существование соседей, если там их нет (клетка у какой-либо границы поля)
'установить в номер того соседа -1
'-----------------------------------------------
If (Tek / x) = (Int(Tek / x)) Then Sosedi(1) = -1: Sosedi(2) = -1: Sosedi(3) = -1
If ((Tek - 1) / x) = (Int((Tek - 1) / x)) Then Sosedi(5) = -1: Sosedi(6) = -1: Sosedi(7) = -1
If (Tek - x) <= 0 Then Sosedi(1) = -1: Sosedi(7) = -1: Sosedi(8) = -1
If (x + Tek) > Maximum Then Sosedi(3) = -1: Sosedi(4) = -1: Sosedi(5) = -1
'Определить номера соседей (в массиве Mines)
'------------------------------------------------
If Sosedi(6) <> -1 Then Sosedi(6) = Tek - 1 'F
If Sosedi(2) <> -1 Then Sosedi(2) = Tek + 1 'B
If Sosedi(8) <> -1 Then Sosedi(8) = Tek - x 'H
If Sosedi(4) <> -1 Then Sosedi(4) = Tek + x 'D
'-----------------------------------------------
If Sosedi(7) <> -1 Then Sosedi(7) = Sosedi(8) - 1 ' G=H-1
If Sosedi(1) <> -1 Then Sosedi(1) = Sosedi(8) + 1 ' G=H+1
If Sosedi(5) <> -1 Then Sosedi(5) = Sosedi(4) - 1 ' E=D+1
If Sosedi(3) <> -1 Then Sosedi(3) = Sosedi(4) + 1 ' E=D+1
'Схема массива для соседних клеток I - клетка для которой определяем соседей

'                  H
'              G 7 8 1 (A)
'              F 6 I 2 (B)
'                5 4 3 (C)
'                E D

End Sub
Sub Inc(ByRef Include As Long, Couter As Long)
' Прибавляет к значению Include значение Couter если оно <> 0, а иначе прибавляет 1
If Couter = 0 Then Include = Include + 1 Else Include = Include + Couter
End Sub
Public Sub GameStart()
'Запуск процесса игры
Form1.Timer1.Enabled = True 'Включаем таймер
Form1.Timer1.Interval = 1000 ' устанавливаем интервал срабатывания в 1 сек.
End Sub
Public Sub SetFlags(flag As Integer)
' Установка флажков
' Если флаг уже был установлен на данной клетке: снимаем его в массиве флагов, присваиваем картинке чека "никакое" изображение (стираем значок флажка), убавляем счетчик флагов, отображаем количество мин незафлагованных
If Flags(flag) Then Flags(flag) = False: Form1.Check1(flag).Picture = InvisItems.None.Picture: FlagCtr = FlagCtr - 1: Form1.Num1.Value = Form1.KolMines - FlagCtr: Exit Sub
' флагов больше чем мин - проверяем, победили или нет
If FlagCtr >= Form1.KolMines Then CheckPobeda: Exit Sub
Flags(flag) = True 'установили флаг в массиве флагов
Form1.Check1(flag).Picture = InvisItems.flag.Picture ' Поменяли картинку на чеке на изображение флажка
FlagCtr = FlagCtr + 1 ' увеличили счетчик флагов
Form1.Num1.Value = Form1.KolMines - FlagCtr ' отобразили количество мин
If FlagCtr = Form1.KolMines Then CheckPobeda: Exit Sub ' проверили, победили или нет

End Sub
Sub CheckPobeda() ' Проверка, победил юзер или флажки кончились
Dim I As Integer
For I = 1 To Form1.MaxFiled 'цикл по количеству клеток на поле
    'если в массиве - мина а флага над ней нет, идем на метку 10
    If (Mines(I) = -1) And (Flags(I) = False) Then GoTo 10
Next I
Pobeda ' иначе победа!
Exit Sub
10 NoMines.Show 1 ' показываем пользователю что эт ему не винда!
End Sub
Public Sub Porajen(Index As Integer)
Dim I As Integer
Form1.mnuGPorajen.Enabled = False 'отключаем возможность сдаться =)
For I = 1 To Form1.MaxFiled
'Мина, на которой мы подорвались. Уст значение чека в checked а картинку - в мину
If I = Index Then Form1.Check1(I).DisabledPicture = InvisItems.mineboom.Picture: Form1.Check1(I).Value = Checked: Form1.Check1(I).Enabled = False: GoTo 10
'Если над миной флаг - просто вырубаем Check
If (Mines(I) = -1) And (Flags(I) = True) Then Form1.Check1(I).DisabledPicture = InvisItems.flag.Picture: Form1.Check1(I).Enabled = False: GoTo 10
'Над миной флага нет, рисуем мину и вырубаем Check
If (Mines(I) = -1) And (Flags(I) = False) Then Form1.Check1(I).DisabledPicture = InvisItems.Mine.Picture: Form1.Check1(I).Value = Checked: Form1.Check1(I).Enabled = False: GoTo 10
'Флаг есть а мины нет! Душибка!!! Рисуем крестег =)
If (Mines(I) <> -1) And (Flags(I) = True) Then Form1.Check1(I).DisabledPicture = InvisItems.errors.Picture: Form1.Check1(I).Value = Checked: Form1.Check1(I).Enabled = False: GoTo 10
Form1.Check1(I).Enabled = False ' Вырубаем клетку
10 Next I
Form1.Command1.Picture = InvisItems.Porajen.Picture ' Картинку на кнопке гл. формы меняем на злобный фейс
Form1.IsKonec = True 'устанавливаем флаг конца игры
Form1.Timer1.Enabled = False ' отключаем таймер
OpenDoor ' открываем дверь CD (легкий инфаркт мимокадра юзеру =)
If FileExist("boom.wav") Then PlayWAVE "boom.wav" ' Проигрываем БАБАХ если файл есть.
CloseDoor 'закрываем CD (слабонервных отвозит скорая помошь)
End Sub
Public Sub OpenNull(Index As Integer)
'открывает свободные клетки
Dim I As Integer 'рабочая
Dim IsZero As Boolean 'флаг пустой клетки
Dim Zero(750) As Long 'сюда записывыем номера клеток которые нашли и надо открывать
Dim SPointer As Integer ' стартовый указатель в массиве Zero
Dim TPointer As Integer ' текущий указатель там же
TPointer = 1 'текущий инициализируем 1
SPointer = 0 'Стартовый указатель - 0
Form1.Check1(Index).Value = Checked ' устанавливаем чек Index - то на что нажали в checked
Form1.Check1(Index).Enabled = False ' отключаем чек
yyy: SPointer = SPointer + 1 ' прибавляем стартовый указатель (в массиве Zero)
GetSosedi CLng(Index), Form1.UserX, Form1.MaxFiled 'Получаем соседей Index не забывая его преобразовать в Long - CLng
For I = 1 To 8
    If Sosedi(I) = -1 Then GoTo 10 ' нет соседней клетки - пошли следующую анализировать
    If Flags(Sosedi(I)) Then GoTo 10 ' в соседней клетке флаг - аналогично
    If Mines(Sosedi(I)) > 0 Then Cifra (Sosedi(I)): GoTo 10 ' в соседней клетке цифра, ставим ее и аналогично
    If Mines(Sosedi(I)) = -1 Then GoTo 10 ' в соседней клетке мина - аналогично
    If Mines(Sosedi(I)) = -2 Then GoTo 10 ' соседняя клетка проанализирована - аналогично
    Zero(TPointer) = Sosedi(I) ' запоминаем номер пустой клетки в массив
    Mines(Sosedi(I)) = -2 ' ставим признак, то что уже проанализировали эту клетку
    TPointer = TPointer + 1 ' текущий указатель увеличиваем
10 Next I
'Если в массиве Zero по стартовому указателю не 0 - значит, мы
'нашли не все клетки которые можно открыть. Присваиваем Index следующую клетку, которую будем анализировать.
' и повторяем процедуру анализа, уходим на yyy
If Zero(SPointer) <> 0 Then Index = Zero(SPointer): GoTo yyy
' открываем чеки на форме
For I = 1 To 750
    If Zero(I) = 0 Then GoTo 20 'Нашли 0 эл-т  массиве Zero - больше открывать нечо, уходим
    Form1.Check1(Zero(I)).Value = Checked ' устанавливаем чек во включенное состояние
    Form1.Check1(Zero(I)).Enabled = False ' отключаем чек
Next I
20
End Sub
Sub Cifra(Index As Integer)
'отображение цифры на соотв чеке на форме
Form1.Check1(Index).DisabledPicture = InvisItems.Numers(Mines(Index)).Picture
' устанавливаем в свойство чека (DisabledPicture - Картинка в отключенном состоянии) в соответствии
' с цифрой для этого чека (из массива Mines) картинку из массива контролов Pic4ureBox Numers
Form1.Check1(Index).Value = Checked ' устанавливаем чек во включенное состояние
Form1.Check1(Index).Enabled = False ' отключаем чек
End Sub
Sub Pobeda()
' действие при победе
Dim I As Integer 'рабочая
For I = 1 To Form1.MaxFiled 'ставим флаги и отключаем чеки
    ' Если есть мина - ставим над ней флаг - устанавливаем признак флага, картинку и отключаем чек
    If Mines(I) = -1 Then Flags(I) = True: Form1.Check1(I).Picture = InvisItems.flag.Picture: _
    Form1.Check1(I).DisabledPicture = InvisItems.flag.Picture
    Form1.Check1(I).Enabled = False
Next I
Form1.IsKonec = True ' признак конца игры
Form1.Command1.Picture = InvisItems.Pobeda.Picture ' на кнопочку на форме вешаем картинку смайлег ф очках
Form1.Timer1.Enabled = False ' отключаем таймер
If FileExist("pobeda.wav") Then PlayWAVE "pobeda.wav" ' если есть звуковой файл - проигрываем
End Sub
