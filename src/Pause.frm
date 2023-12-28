VERSION 5.00
Begin VB.Form Pause 
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   Picture         =   "Pause.frx":0000
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Конец паузы  - клавиша ESC или щелчок мышью"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Pause"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Click()
'обработчик клика по форме
Unload Me 'выгрузка объекта формы
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
' обработчик нажатия клавиши
If KeyAscii = 27 Then Unload Me 'Если нажата ESC - выгрузка объекта формы
End Sub
Private Sub Form_Load()
'загрузка формы
Form1.Timer1.Enabled = False 'выключаем таймер
'подстраиваемся своими размерами под размеры главного окна
Me.Top = Form1.Top
Me.Left = Form1.Left
Me.Width = Form1.Width
Me.Height = Form1.Height
End Sub
Private Sub Form_Unload(Cancel As Integer)
' событие выгрузки формы
Form1.Timer1.Enabled = True 'включаем таймер
End Sub
Private Sub Label1_Click()
'обработчик клика по надписи
Unload Me
End Sub

