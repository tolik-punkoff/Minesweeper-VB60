VERSION 5.00
Begin VB.Form SaveFile 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Сохранить"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4620
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4620
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Закрыть"
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Удалить"
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Сохранить"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Введите или выберите из списка имя файла"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "SaveFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
Dim Otv As VbMsgBoxResult
If Text1.Text = "" Then Exit Sub
If Text1.Text = ".ssg" Then Exit Sub
Otv = vbYes
If GetExt(Text1.Text) <> ".ssg" Then Text1.Text = Text1.Text + ".ssg"
If Not ExistInList(Text1.Text, List1) Then List1.AddItem (Text1.Text)
If FileExist(Text1.Text) Then Otv = MsgBox("Заменить файл?", vbQuestion + vbYesNo, "Внимание!")
If Otv <> vbYes Then Exit Sub
SaveFileSSF Text1.Text
Unload Me
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Command2_Click()
On Error GoTo 10
Dim Otv As VbMsgBoxResult
Otv = MsgBox("Удалить файл?", vbExclamation + vbYesNo, "Внимание!")
If Otv = vbYes Then Kill Text1.Text: Form_Load: Text1.Text = ""
Exit Sub
10 MsgBox Err.Source + ": " + Err.Description + vbCrLf + vbCrLf + "#" + Str(Err.Number), vbCritical, "Ошибка файла"
End Sub

Private Sub Command2_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command3_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
Form1.Timer1.Enabled = False
List1.Clear
CreateList "*.ssg", List1
End Sub
Private Sub Form_Unload(Cancel As Integer)
Form1.Timer1.Enabled = True
End Sub

Private Sub List1_Click()
Text1.Text = List1.List(List1.ListIndex)
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub
