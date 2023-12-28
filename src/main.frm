VERSION 5.00
Object = "{B655A5F2-4B41-11D3-9C70-00C058205D4C}#1.0#0"; "INDCTR.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�������"
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
      ToolTipText     =   "����� ��-�����"
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
      Caption         =   "����"
      Begin VB.Menu mnuGNew 
         Caption         =   "�����"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuGPorajen 
         Caption         =   "�������"
      End
      Begin VB.Menu mnuGPause 
         Caption         =   "�����"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuGExit 
         Caption         =   "����� "
      End
   End
   Begin VB.Menu mnuUroven 
      Caption         =   "�������"
      Begin VB.Menu mnuUChai 
         Caption         =   "������"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuUBil 
         Caption         =   "�������"
      End
      Begin VB.Menu mnuUKr 
         Caption         =   "������"
      End
      Begin VB.Menu MnuUDr 
         Caption         =   "������..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "������"
      Begin VB.Menu mnuHAbout 
         Caption         =   "� ���������"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Width - ������
'Height - ������
Option Explicit
Public KolMines, MaxFiled As Long, OldFiled As Long ' ���������� ���, ������������ ����������� ����, ���������� ��� �������� ���������� ����������� ���� (��������)
Public UserX As Long, UserY As Long ' ����������� ���� �� ������, ������
Public IsStart As Boolean ' ������� ������ ����
Public IsKonec As Boolean ' ������� ����� ����
Public DownIdx As Integer ' ������ ����, �� ������� ��������
Public Tick As Integer ' ���������� ���������� ������� � ������ ����
Public I As Integer ' ������� ����������
Public OpnCtr As Long ' ���������� �������� ������
Sub NewMineArray(KolMines, NMax As Long)
' ��������� �������� ������� ��� NewMineArray(����������_���, ����������_������_���� As Long)
Dim I, Z As Long, Coeft, J As Integer
FlagCtr = 0 ' �������� ���������� ������������� �������
' ������� ��� ���� � ��� ������ (������� ��� � ������� ����������� - �� 750 ��-���)
' ����� �������� ���� ����������� �������� � ������ ������ ���������
For I = 1 To 750 ' ����� ����� �������
    Mines(I) = 0 ' ������� ��-�� ������� ���
    Flags(I) = False '������� ��-�� ������� �������
Next I ' ����� ����� �������
' �������� �����������, �� ������� ����� ��������� ��������� �����
' � ����������� �� ���������� ������ ����
If NMax <= 10 Then Coeft = 10: GoTo Cycle
If NMax <= 100 Then Coeft = 100: GoTo Cycle
If NMax <= 1000 Then Coeft = 1000: GoTo Cycle
Cycle:
For I = 1 To KolMines '������������� ���� (���� �� 1 �� ���������� ���)
Randomize Timer '���������� ��������� ���������� ��������� �����
10 Z = Int(Rnd(NMax) * Coeft) '�������� ����� ������ � ������� ����� ������������� ����
   '����� ��������� ����� � ����������� �� ���-�� ������ �� ����
   ' ��������� �� ����������� (��. ����)
   ' � ����� ����� ����� - � Z - ����� �������������� ������ ���� (����� �� ����� � ������� ���)
If Z > NMax Then Z = Int(Z / 10): GoTo 10 ' ���� Z >NMax              |
20 If Mines(Z) = -1 Then GoTo 10 ' ���� ���� � ������ ��� ����������� | - ��������� ��������� ������
If Z = 0 Then GoTo 10 ' ������� ��-� �� ������������                  |
Mines(Z) = -1 ' ������ �������� - ������ ������ ������� ���� (-1)

'------------------------------------------------------------------------------
'��������� ������� ����
GetSosedi Z, UserX, NMax '�������� ������� ������. ��������� ���������� ������ �������� ������ �  ������ sosedi
For J = 1 To 8 ' ���� �� ���-�� ������� - � ������ ������ �/� �������� 8 ��������
    If Sosedi(J) = -1 Then GoTo Net ' ��� �������
    If Mines(Sosedi(J)) = -1 Then GoTo Net '����� - ����
    Inc Mines(Sosedi(J)), 0 ' ����� ������ ��� ����� - +1 (� ������� Mines ����� �������� ���-� � ���-�� ��� � �������� �������)
Net: Next J ' ����� ����� �� �������
Next I ' ����� ����� ��������� ���
End Sub
Sub DeleteFiled(Nach, Konez As Long)
' ��������� �������� �������� (�����) � �������� ����
' DeleteFiled(1_���������_��-�, ���������_���������_��-�)
On Error GoTo 20 ' � ������ ������ ���� �� ����� 20
Dim I As Integer ' ������� �����
Check1(1).Enabled = True ' �������� 1-� ���
Check1(1).Picture = InvisItems.None.Picture ' ��������� ��� ������ ��������
Check1(1).Value = Unchecked ' ���������� ��� �������� � "��������"
For I = Nach To Konez '������� ����
    Unload Check1(I) ' ��������� ������ �� ������
Next I ' ����� ���. �����
Exit Sub ' ������� �� ���������
20 MsgBox Err.Description ' ��������� �� ������
End Sub
Sub CreateFiled() ' ��������� �������� �������� ����
Dim I, Z As Integer
DeleteFiled 2, OldFiled ' ������� ������ ������� ����
Z = 1 ' ��������� ���� � "������" ����
' ��������� ������/������ ������� ����� � ����������� �� ���-��
' �����-������ �� ���������/����������� ������������ 100 � 820 �����������, ���������
' ������� �������� ���� � �������� ������� =)
Form1.Width = Check1(1).Width * UserX + 100
Form1.Height = Check1(1).Height * UserY + Line1.Y1 + 820
For I = 2 To MaxFiled ' ���� ��������� ����� �� 2-�� (1-� ����, ������ ����������) �� ���������� ������ �� ����
    Load Check1(I) ' ��������� ��� � ������ ���������� ���� �����
    Check1(I).Visible = True ' ������ ��� �������
    Check1(I).Left = Check1(I - 1).Left + Check1(1).Width ' ������������� ��� ���������� �� ����� � ����������� �� �������� ������ ���� � ��������� ��������� � ��� ����
    Check1(I).Top = Check1(I - 1).Top
    Z = Z + 1 ' ���������� 1 � ��������� ���� � ������ ����
    If Z > UserX Then Check1(I).Top = Check1(I - 1).Top + Check1(1).Height: Check1(I).Left = 0: Z = 1
    ' ���� ������� ���������, ���������� ��� ����, � ����������� Z 1-��, �.�. ���� ��� ����� 1-� � ����� ������
Next I
End Sub
Private Sub Check1_Click(Index As Integer)
' ��������� ����� �� ����
Static NFirstStart As Boolean ' ����, ����������, ��� ��� �� 1-� ������� �� ��� ����
' ���� �� ������������� ���-�� �������� ������ � 1 � ��� ���� � TRUE
If NFirstStart = False Then OpnCtr = 1: NFirstStart = True
If Flags(Index) Then Check1(Index).Value = Unchecked: Exit Sub '���� �� ���� ����� ��� ���������� ����, ������ ���, ��� ������� �� ���� =) (unchecked) � ������� �� ����������� �������
'---------------------------------------------------------------
If Mines(Index) <> -1 Then OpnCtr = OpnCtr + 1 Else: DownIdx = Index: Exit Sub ' ���� ��� �� ���� (� ����� � �������� ���) ���������� ���-�� �������� ������ � ����������, ���� ���� - ��������� ����� ������� ������ � ������
If OpnCtr > MaxFiled - KolMines Then Pobeda: ' ���� ���-�� �������� ������ ������ ��� MaxFiled - KolMines ��������� ��������� ������.
DownIdx = Index '��������� ����� ������� ������
End Sub
Sub Check1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
' ��������� ������� ������� �� ������ ���� �� ����
If Not IsStart Then GameStart ' ���� ���� �� ������ - ������ ����
If Button = 2 Then SetFlags Index: Exit Sub ' ���� ������ ������ ������ ���� - ������������� ���� � ������� �� ����������� �������
Command1.Picture = InvisItems.O.Picture '������ �������� �� ������ ������ ���� �� "����������" �����
End Sub
Private Sub Check1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
' ��������� ������� ���������� ������ ���� �� ����
If Not IsKonec Then Command1.Picture = InvisItems.Smile.Picture '���� �� ����� ������ �������� �� ��������
If DownIdx <> Index Then Exit Sub '������� ����� ��������� ����� ������ ����, ������� �� �����������
'���� � ������ ����� - ������ ��
If Mines(Index) > 0 Then Cifra (Index)
'----------------------------------------------------------------
'� ���� ���� :-))) ��������� ��������� ���������
If Mines(Index) = -1 Then Porajen Index: Exit Sub
'----------------------------------------------------------------
'����� - ������������� ��� � ������� ���������, ��������� ���, ��������� �������� ������ ������
If Mines(Index) = 0 Then Check1(Index).Value = Checked: Check1(Index).Enabled = False: OpenNull (Index): Exit Sub
End Sub
Private Sub Command1_Click()
'���������� ������� ������� �� ������ � ������� ����
Timer1.Interval = 0 ' �������� �������� ������������ �������
Timer1.Enabled = False '��������� ������
Num2.Value = 0 ' �������� �����
Timer1.Enabled = True ' �������� ������
' � ����������� �� ���������� ������ ��������� ��������� ���������������
' ���������� ������� (���� ���� �� ��������������� ���� � ������� ���������)
If mnuUBil.Checked Then mnuUBil_Click: Exit Sub
If mnuUChai.Checked Then mnuUChai_Click: Exit Sub
If MnuUDr.Checked Then Form_Load: Exit Sub
If mnuUKr.Checked Then mnuUKr_Click: Exit Sub
End Sub

Private Sub Form_Load()
Me.MousePointer = 11 ' ������������� ������ � ���� �������
Tick = 0 ' �������� �����
mnuGPorajen.Enabled = True ' �������� ����������� �������
' ������������� �������� ��� ���������
' (�� ��� ��������� ������� ����), �������� �������� � �����-
' ���������� InvisItems
Check1(1).Picture = InvisItems.None.Picture '�� ���������� ��������� - ������ ��������
Check1(1).DisabledPicture = InvisItems.None.Picture '� ����������� ��������� - ������ ��������
Check1(1).Enabled = True '���������� ���
Check1(1).Visible = True '������ ��� �������
Command1.Picture = InvisItems.Smile.Picture '�� ������ ������� ���� - �������� � ���������
'������������� ����� �����/������ ����
'���y���� ������� � �������� ������� � "������������ �������"
IsStart = False '���� ������ ���� - � ����
IsKonec = False '���� ����� ���� - � ����
OpnCtr = 0 '���������� �������� �������� ���� ��������
I = 0: Timer1.Enabled = False: Num2.Value = 0 '�������� ��������, ��������� ������, �������� �-�� �������� ���
If UserX <= 0 Then UserX = 10 ' ��������� �������� ����������� �������� ����
If UserY <= 0 Then UserY = 10:
If KolMines <= 0 Then KolMines = 10 '���� ���������� ��� ���� ������ ����������� - ������������� �� � 10
OldFiled = MaxFiled ' ��������� ������ �������� ����������� ����
MaxFiled = UserX * UserY     ' �������� ����
NewMineArray KolMines, MaxFiled '���������� ����
CreateFiled '������� ����
Form1.Num1.Value = Form1.KolMines ' ������� ���������� ��� �� ����� � ������� � ��. �������
Me.MousePointer = 0 ' ������������� ���������� ������
End Sub
Private Sub Form_Resize()
' ���������� ������� ��������� ������� ������
Dim Kn As Long
Num1.Top = 0 ' ����������� ������ ���������
Num1.Left = 120
Line1.X2 = Form1.ScaleWidth ' ��������� �����
Num2.Top = 0 ' ����������� 2�
Num2.Left = Form1.ScaleWidth - 850
'����������� ������
Command1.Left = Num1.Left + Num1.Height + 450
Kn = Num2.Left - Command1.Left
Kn = Kn - Command1.Height - Command1.Height - Command1.Height
Kn = Int(Kn / 2) + Num1.Left + Num1.Height + 1020
Command1.Left = Kn
' ����� - ����������� ������������. ������� ������� �������� ���� � �������� �������
End Sub
Private Sub Form_Unload(Cancel As Integer)
' ��������� ������� ������ (������ �� �������)
Dim Z As VbMsgBoxResult ' ������� ���������� ��� ��������� messagebox �
Z = MsgBox("������ ����?", vbQuestion + vbYesNo, "�����")
If Z = vbYes Then End ' � ���� ����� �� �� ����� ���������
'� ��������� ������ �������� �������� �����
Cancel = 1
End Sub
Private Sub mnuGExit_Click()
'����� ����� ����
Form_Unload (0)
End Sub
Private Sub mnuGPorajen_Click()
' ���������� ������ ���� "�������"
Dim Otv As VbMsgBoxResult
' ����������, ����� �� �������.
Otv = MsgBox("�� ������������� ������ �������???", vbYesNo + vbQuestion + vbSystemModal, "�������")
' ��� - ������� �� �����������
If Otv = vbNo Then Exit Sub
' ���� ����
For I = 1 To 750
If Mines(I) = -1 Then Porajen I:  Exit Sub '� ��� ������ ���� �������� ��������� ���������
Next I
End Sub
Private Sub mnuGNew_Click()
' ���������� ������ ���� "����� ����"
Command1_Click '����� ����
End Sub
Private Sub mnuGPause_Click()
' ���������� ������ ���� "�����"
Pause.Show 1 ' ����� - ������� ����� �����
End Sub

Private Sub mnuHAbout_Click()
' ���������� ������ ���� "� ���������"
frmAbout.Show 1 ' ������� ����� � ��������
End Sub

Private Sub mnuUBil_Click()
'��������� ������ "�������"
mnuUChai.Checked = False ' ������� ��� � ��������� �������
MnuUDr.Checked = False
mnuUKr.Checked = False
mnuUBil.Checked = True ' ������������� ��� �� �������
UserX = BivalX ' ����������� ���� ����������� ���������� ��� ���������������� ������
UserY = BivalY
KolMines = BivalMines ' ���������� ��� = ���������� ��� ���������������� ������
Form_Load ' ������������� ������� ����
End Sub

Private Sub mnuUChai_Click()
'��������� ������ "������"
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
'��������� ������ "������"
mnuUChai.Checked = False
MnuUDr.Checked = True
mnuUKr.Checked = False
mnuUBil.Checked = False
'����� ����� � ������� �������� ����������� ����
Form2.Show 1
'���� ������������ ������� ����� �� ������ �� ���������
If Form2.GetCancel Then Exit Sub
UserX = Form2.OthX
UserY = Form2.OthY
KolMines = Form2.OthMines
Form_Load
End Sub

Private Sub mnuUKr_Click()
'��������� ������ "������"
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
' ��������� ������� ������� - ������ ������� ���������� �����
' � ������� ��� � ������� � �������
Tick = Tick + 1
Num2.Value = Tick
If Tick = 999 Then Tick = 0 '��� ���� ������� �� ������� - � ���� ������ �� ����� ����
End Sub
