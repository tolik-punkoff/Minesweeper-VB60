VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Размер игрового поля"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3765
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   3765
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Отмена"
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox M 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Text            =   "0"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox V 
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Text            =   "0"
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox H 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Text            =   "0"
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Количество мин"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Ширина"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Высота"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'H - Высота (Y) V - Длина (X)
Public OthX, OthY, OthMines As Integer
Public GetCancel As Boolean
Private Sub Command1_Click()
OthY = Val(H.Text) ' из текстовых контролов берем данные, переводим в числа и
OthX = Val(V.Text) ' засовываем их в переменные
OthMines = Val(M.Text)
If Val(M.Text) > OthX * OthY Then OthMines = 10
If Val(H.Text) < 10 Then OthY = 10 ' Проверка и установка значений
If Val(H.Text) > 30 Then OthY = 25 'чтоб не было слишком большое/маленькое поле
If Val(V.Text) < 10 Then OthX = 10 ' и слишком много/мало мин
If Val(V.Text) > 30 Then OthX = 30
If Val(M.Text) < 10 Then OthMines = 10
If Val(M.Text) > 666 Then OthMines = 10
GetCancel = False ' признак того, что пользователь отменил вызов формы устанавливаем в FALSE
Unload Me 'выгрузка объекта формы
End Sub
Private Sub Command2_Click()
GetCancel = True ' признак того, что пользователь отменил вызов формы устанавливаем в TRUE
Form1.Timer1.Enabled = True ' включаем таймер
Unload Me 'выгрузка объекта формы
End Sub
Private Sub Form_Load()
Form1.Timer1.Enabled = False ' выключаем таймер
M.Text = Str(Form1.KolMines) ' загружаем в текстовые поля текущие значения
H.Text = Str(Form1.UserY) ' размерности поля и мин, переводим перед этим их в строки
V.Text = Str(Form1.UserX)
End Sub
