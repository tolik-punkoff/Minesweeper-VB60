VERSION 5.00
Object = "{B655A5F2-4B41-11D3-9C70-00C058205D4C}#1.0#0"; "INDCTR.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Саперег"
   ClientHeight    =   3195
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3930
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   2400
      Top             =   -120
   End
   Begin Indctr.Num Num2 
      Height          =   390
      Left            =   3120
      TabIndex        =   1
      Top             =   0
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   688
      BackColor       =   12632256
      NumColor        =   16776960
      Max             =   999
   End
   Begin VB.CommandButton Command1 
      Default         =   -1  'True
      Height          =   615
      Left            =   1680
      Picture         =   "main.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Давай по-новой"
      Top             =   0
      Width           =   615
   End
   Begin Indctr.Num Num1 
      Height          =   390
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   688
      Max             =   999
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Index           =   1
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderStyle     =   6  'Inside Solid
      X1              =   0
      X2              =   5040
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Menu mnuGame 
      Caption         =   "Игра"
      Begin VB.Menu mnuGNew 
         Caption         =   "Новая"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuGPorajen 
         Caption         =   "Сдаться"
      End
      Begin VB.Menu mnuGPause 
         Caption         =   "Пауза"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuGExit 
         Caption         =   "Выход "
      End
   End
   Begin VB.Menu mnuUroven 
      Caption         =   "Уровень"
      Begin VB.Menu mnuUChai 
         Caption         =   "Чайник"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuUBil 
         Caption         =   "Бывалый"
      End
      Begin VB.Menu mnuUKr 
         Caption         =   "Крутой"
      End
      Begin VB.Menu MnuUDr 
         Caption         =   "Другой..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Помощь"
      Begin VB.Menu mnuHAbout 
         Caption         =   "О программе"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Width - ширина
'Height - высота
Option Explicit
Public KolMines, MaxFiled As Long, OldFiled As Long ' Количество мин, максимальная размерность поля, переменная для хранения предыдущей размерности поля (буферная)
Public UserX As Long, UserY As Long ' размерность поля по высоте, ширине
Public IsStart As Boolean ' признак начала игры
Public IsKonec As Boolean ' признак конца игры
Public DownIdx As Integer ' индекс чека, на который кликнули
Public Tick As Integer ' количество прошедшего времени с начала игры
Public I As Integer ' рабочая переменная
Public OpnCtr As Long ' количество открытых клеток
Sub NewMineArray(KolMines, NMax As Long)
' процедура создания массива мин NewMineArray(Количество_мин, Количество_клеток_поля As Long)
Dim I, Z As Long, Coeft, J As Integer
FlagCtr = 0 ' обнуляем количество установленных флажков
' очищаем все мины и все флажки (массивы мин и флажков статические - по 750 эл-тов)
' таких размеров поле максимально возможно в данной версии программы
For I = 1 To 750 ' старт цикла очистки
    Mines(I) = 0 ' очистка эл-та массива мин
    Flags(I) = False 'очистка эл-та массива флажков
Next I ' конец цикла очистки
' выбираем коэффициент, на который будем домножать случайное число
' в зависимости от количества клеток поля
If NMax <= 10 Then Coeft = 10: GoTo Cycle
If NMax <= 100 Then Coeft = 100: GoTo Cycle
If NMax <= 1000 Then Coeft = 1000: GoTo Cycle
Cycle:
For I = 1 To KolMines 'устанавливаем мины (цикл от 1 до количества мин)
Randomize Timer 'сбрасываем установки генератора случайных чисел
10 Z = Int(Rnd(NMax) * Coeft) 'получаем номер клетки в которую будем устанавливать мину
   'берем случайное число в зависимости от кол-ва клеток на поле
   ' домножаем на коэффициент (см. выше)
   ' и берем целую часть - в Z - номер обрабатываемой клетки поля (соотв ей номер в массиве мин)
If Z > NMax Then Z = Int(Z / 10): GoTo 10 ' Если Z >NMax              |
20 If Mines(Z) = -1 Then GoTo 10 ' Если мина в клетке уже установлена | - повторяем процедуру выбора
If Z = 0 Then GoTo 10 ' Нулевой эл-т не используется                  |
Mines(Z) = -1 ' клетка свободна - значит ставим признак мины (-1)

'------------------------------------------------------------------------------
'Установка соседей мины
GetSosedi Z, UserX, NMax 'получаем соседей клетки. процедура сбрасывает номера соседних клеток в  массив sosedi
For J = 1 To 8 ' цикл по кол-ву соседей - у каждой клетки м/б максимум 8 соседних
    If Sosedi(J) = -1 Then GoTo Net ' Нет соседей
    If Mines(Sosedi(J)) = -1 Then GoTo Net 'Сосед - мина
    Inc Mines(Sosedi(J)), 0 ' Сосед пустой или цифра - +1 (в массиве Mines также хранится инф-я о кол-ве мин в соседних клетках)
Net: Next J ' конец цикла по соседям
Next I ' конец цикла установки мин
End Sub
Sub DeleteFiled(Nach, Konez As Long)
' процедура удаления объектов (чеков) с игрового поля
' DeleteFiled(1_удаляемый_эл-т, последний_удаляемый_эл-т)
On Error GoTo 20 ' в случае ошибки идем на метку 20
Dim I As Integer ' счетчик цикла
Check1(1).Enabled = True ' включаем 1-й чек
Check1(1).Picture = InvisItems.None.Picture ' загружаем ему пустую картинку
Check1(1).Value = Unchecked ' выставляем его значение в "невыбран"
For I = Nach To Konez 'рабочий цикл
    Unload Check1(I) ' выгружаем объект из памяти
Next I ' конец раб. цикла
Exit Sub ' выходим из процедуры
20 MsgBox Err.Description ' сообщение об ошибке
End Sub
Sub CreateFiled() ' процедура создания игрового поля
Dim I, Z As Integer
DeleteFiled 2, OldFiled ' удаляем старое игровое поле
Z = 1 ' положение чека в "строке" поля
' установка высоты/ширины главной формы в зависимости от кол-ва
' чеков-клеток по вертикали/горизонтали Коэффициенты 100 и 820 поправочные, подобраны
' методом научного тыка в процессе отладки =)
Form1.Width = Check1(1).Width * UserX + 100
Form1.Height = Check1(1).Height * UserY + Line1.Y1 + 820
For I = 2 To MaxFiled ' цикл установки чеков от 2-го (1-й есть, создан статически) до количества клеток на поле
    Load Check1(I) ' добавляем чек в массив управления этих чеков
    Check1(I).Visible = True ' делаем его видимым
    Check1(I).Left = Check1(I - 1).Left + Check1(1).Width ' устанавливаем его координаты на форме в зависимости от размеров самого чека и положения соседнего с ним чека
    Check1(I).Top = Check1(I - 1).Top
    Z = Z + 1 ' прибавляем 1 к положению чека в строке поля
    If Z > UserX Then Check1(I).Top = Check1(I - 1).Top + Check1(1).Height: Check1(I).Left = 0: Z = 1
    ' если строчка кончилась, перемещаем чек ниже, и присваиваем Z 1-цу, т.к. этот чек будет 1-м в новой строке
Next I
End Sub
Private Sub Check1_Click(Index As Integer)
' обработка клика по чеку
Static NFirstStart As Boolean ' флаг, провеяющий, что это НЕ 1-е нажатие за всю игру
' Если ДА устанавливаем кол-во открытых клеток в 1 а сам флаг в TRUE
If NFirstStart = False Then OpnCtr = 1: NFirstStart = True
If Flags(Index) Then Check1(Index).Value = Unchecked: Exit Sub 'если на этом месте был установлен флаг, делаем вид, что нажатия не было =) (unchecked) и выходим из обработчика события
'---------------------------------------------------------------
If Mines(Index) <> -1 Then OpnCtr = OpnCtr + 1 Else: DownIdx = Index: Exit Sub ' если там НЕ мина (в соотв с массивом мин) прибавляем кол-во открытых клеток и продолжаем, если мина - указываем номер нажатой клетки и уходим
If OpnCtr > MaxFiled - KolMines Then Pobeda: ' если кол-во открытых клеток больше чем MaxFiled - KolMines запускаем процедуру победы.
DownIdx = Index 'указываем номер нажатой клетки
End Sub
Sub Check1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
' Обработка события нажатия на кнопку мыши на чеке
If Not IsStart Then GameStart ' если игра не начата - начать игру
If Button = 2 Then SetFlags Index: Exit Sub ' если нажата правая кнопка мыши - устанавливаем флаг и выходим из обработчика события
Command1.Picture = InvisItems.O.Picture 'Меняем картинку на кнопке старта игры на "испуганный" смайл
End Sub
Private Sub Check1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
' Обработка события отпускания кнопки мыши на чеке
If Not IsKonec Then Command1.Picture = InvisItems.Smile.Picture 'Если не конец меняем картинку на улыбочку
If DownIdx <> Index Then Exit Sub 'условие будет выполнено когда ставим флаг, выходим из обработчика
'Если в клетке цифра - ставим ее
If Mines(Index) > 0 Then Cifra (Index)
'----------------------------------------------------------------
'А если мина :-))) Запускаем процедуру поражения
If Mines(Index) = -1 Then Porajen Index: Exit Sub
'----------------------------------------------------------------
'Пусто - устанавливаем чек в нажатое состояние, отключаем его, открываем соседние пустые клетки
If Mines(Index) = 0 Then Check1(Index).Value = Checked: Check1(Index).Enabled = False: OpenNull (Index): Exit Sub
End Sub
Private Sub Command1_Click()
'обработчик события нажатия на кнопку в главном окне
Timer1.Interval = 0 ' обнуляем интервал срабатывания таймера
Timer1.Enabled = False 'выключаем таймер
Num2.Value = 0 ' обнуляем время
Timer1.Enabled = True ' включаем таймер
' в зависимости от выбранного уровня сложности запускаем соответствующий
' обработчик события (типа клик на соответствующее меню с уровнем сложности)
If mnuUBil.Checked Then mnuUBil_Click: Exit Sub
If mnuUChai.Checked Then mnuUChai_Click: Exit Sub
If MnuUDr.Checked Then Form_Load: Exit Sub
If mnuUKr.Checked Then mnuUKr_Click: Exit Sub
End Sub

Private Sub Form_Load()
Me.MousePointer = 11 ' устанавливаем курсор в виде часиков
Tick = 0 ' обнуляем время
mnuGPorajen.Enabled = True ' включаем возможность здаться
' устанавливаем картинки для чекбоксов
' (на них построено игровое поле), картинки хранятся в форме-
' контейнере InvisItems
Check1(1).Picture = InvisItems.None.Picture 'во включенном состоянии - пустая картинка
Check1(1).DisabledPicture = InvisItems.None.Picture 'в отключенном состоянии - пустая картинка
Check1(1).Enabled = True 'активируем чек
Check1(1).Visible = True 'делаем его видимым
Command1.Picture = InvisItems.Smile.Picture 'на кнопку запуска игры - картинку с улыбочкой
'устанавливаем флаги конца/старта игры
'обнyляем таймеры и обнуляем контрол с "электронными цифрами"
IsStart = False 'флаг начала игры - в ЛОЖЬ
IsKonec = False 'флаг конца игры - в ЛОЖЬ
OpnCtr = 0 'количество открытых клеточек поля обнуляем
I = 0: Timer1.Enabled = False: Num2.Value = 0 'обнуляем счетчики, отключаем таймер, обнуляем к-во открытых мин
If UserX <= 0 Then UserX = 10 ' Начальные значения размерности игрового поля
If UserY <= 0 Then UserY = 10:
If KolMines <= 0 Then KolMines = 10 'если количество мин было задано неправильно - устанавливаем их в 10
OldFiled = MaxFiled ' сохрвняем старое значение размерности поля
MaxFiled = UserX * UserY     ' Получить поле
NewMineArray KolMines, MaxFiled 'Установить мины
CreateFiled 'создаем поле
Form1.Num1.Value = Form1.KolMines ' выводим количество мин на форму в контрол с эл. цифрами
Me.MousePointer = 0 ' устанавливаем нормальный курсор
End Sub
Private Sub Form_Resize()
' обработчик события изменения размера формой
Dim Kn As Long
Num1.Top = 0 ' Переместили первый индикатор
Num1.Left = 120
Line1.X2 = Form1.ScaleWidth ' Расширили линию
Num2.Top = 0 ' Переместили 2й
Num2.Left = Form1.ScaleWidth - 850
'переместили кнопку
Command1.Left = Num1.Left + Num1.Height + 450
Kn = Num2.Left - Command1.Left
Kn = Kn - Command1.Height - Command1.Height - Command1.Height
Kn = Int(Kn / 2) + Num1.Left + Num1.Height + 1020
Command1.Left = Kn
' числа - поправочные коэффициенты. Выбраны методом научного тыка в процессе отладки
End Sub
Private Sub Form_Unload(Cancel As Integer)
' обработка события выхода (нажатиё на крестик)
Dim Z As VbMsgBoxResult ' заводим переменную под результат messagebox а
Z = MsgBox("Хотите уйти?", vbQuestion + vbYesNo, "Выход")
If Z = vbYes Then End ' и если ответ да то кинец программы
'в противном случае отменяем выгрузку формы
Cancel = 1
End Sub
Private Sub mnuGExit_Click()
'Выход через меню
Form_Unload (0)
End Sub
Private Sub mnuGPorajen_Click()
' обработчик пункта меню "Сдаться"
Dim Otv As VbMsgBoxResult
' спрашиваем, хотят ли сдаться.
Otv = MsgBox("Вы действительно хотите сдаться???", vbYesNo + vbQuestion + vbSystemModal, "СДАТЬСЯ")
' нет - выходим из обработчика
If Otv = vbNo Then Exit Sub
' ищем мину
For I = 1 To 750
If Mines(I) = -1 Then Porajen I:  Exit Sub 'и для первой мины вызываем процедуру поражения
Next I
End Sub
Private Sub mnuGNew_Click()
' обработчик пункта меню "Новая игра"
Command1_Click 'новая игра
End Sub
Private Sub mnuGPause_Click()
' обработчик пункта меню "Пауза"
Pause.Show 1 ' пауза - выводим форму паузы
End Sub

Private Sub mnuHAbout_Click()
' обработчик пункта меню "О программе"
frmAbout.Show 1 ' выводим форму О пограмме
End Sub

Private Sub mnuUBil_Click()
'установка режима "бывалый"
mnuUChai.Checked = False ' снимаем чек с остальных режимов
MnuUDr.Checked = False
mnuUKr.Checked = False
mnuUBil.Checked = True ' устанавливаем чек на бывалом
UserX = BivalX ' Размерности поля присваеваем координаты для соответствующего режима
UserY = BivalY
KolMines = BivalMines ' количеству мин = количество мин соответствующего режима
Form_Load ' перезагружаем игровое поле
End Sub

Private Sub mnuUChai_Click()
'установка режима "чайник"
mnuUChai.Checked = True
MnuUDr.Checked = False
mnuUKr.Checked = False
mnuUBil.Checked = False
UserX = ChainX
UserY = ChainY
KolMines = ChainMines
Form_Load
End Sub
Private Sub MnuUDr_Click()
'установка режима "другой"
mnuUChai.Checked = False
MnuUDr.Checked = True
mnuUKr.Checked = False
mnuUBil.Checked = False
'вызов формы в которой задоется размерность поля
Form2.Show 1
'если пользователь отменил форму то уходим из процедуры
If Form2.GetCancel Then Exit Sub
UserX = Form2.OthX
UserY = Form2.OthY
KolMines = Form2.OthMines
Form_Load
End Sub

Private Sub mnuUKr_Click()
'установка режима "крутой"
mnuUChai.Checked = False
MnuUDr.Checked = False
mnuUKr.Checked = True
mnuUBil.Checked = False
UserX = CoolX
UserY = CoolY
KolMines = CoolMines
Form_Load
End Sub
Private Sub Timer1_Timer()
' обработка события таймера - каждую секунду прибавляем время
' и выводим его в контрол с цифрами
Tick = Tick + 1
Num2.Value = Tick
If Tick = 999 Then Tick = 0 'это чтоб контрол не глюкнул - в него больше не лезет цифр
End Sub
