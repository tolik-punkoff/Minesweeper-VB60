VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������ �������� ����"
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
      Caption         =   "&������"
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
      Caption         =   "���������� ���"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "������"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "������"
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
'H - ������ (Y) V - ����� (X)
Public OthX, OthY, OthMines As Integer
Public GetCancel As Boolean
Private Sub Command1_Click()
OthY = Val(H.Text) ' �� ��������� ��������� ����� ������, ��������� � ����� �
OthX = Val(V.Text) ' ���������� �� � ����������
OthMines = Val(M.Text)
If Val(M.Text) > OthX * OthY Then OthMines = 10
If Val(H.Text) < 10 Then OthY = 10 ' �������� � ��������� ��������
If Val(H.Text) > 30 Then OthY = 25 '���� �� ���� ������� �������/��������� ����
If Val(V.Text) < 10 Then OthX = 10 ' � ������� �����/���� ���
If Val(V.Text) > 30 Then OthX = 30
If Val(M.Text) < 10 Then OthMines = 10
If Val(M.Text) > 666 Then OthMines = 10
GetCancel = False ' ������� ����, ��� ������������ ������� ����� ����� ������������� � FALSE
Unload Me '�������� ������� �����
End Sub
Private Sub Command2_Click()
GetCancel = True ' ������� ����, ��� ������������ ������� ����� ����� ������������� � TRUE
Form1.Timer1.Enabled = True ' �������� ������
Unload Me '�������� ������� �����
End Sub
Private Sub Form_Load()
Form1.Timer1.Enabled = False ' ��������� ������
M.Text = Str(Form1.KolMines) ' ��������� � ��������� ���� ������� ��������
H.Text = Str(Form1.UserY) ' ����������� ���� � ���, ��������� ����� ���� �� � ������
V.Text = Str(Form1.UserX)
End Sub
