Attribute VB_Name = "saper0"
Option Explicit '�������� ���������� ����������
' ��������� ������ ������
Public Const ChainMines = 10 '����
Public Const ChainX = 10 '������ �� �����������
Public Const ChainY = 10 '������ �� ���������
' ��������� ������ �������
Public Const BivalMines = 40 '���������� ����
Public Const BivalX = 16
Public Const BivalY = 16
' ��������� ������ ������
Public Const CoolMines = 99 '���������� ����
Public Const CoolX = 30
Public Const CoolY = 16
' ���������� ������ ������
Public OthMines As Long '���������� ����
Public OthX As Long
Public OthY As Long

Public Mines(750) As Long ' ������ � ���������� ��� ���� � ��� ����� ����� � �������� � ������ �������
Public Flags(750) As Boolean ' ������ � ���������� ��� �����
Public Sosedi(8) As Long ' ������ � �������� �������� ������ (�. ������� GetSosedi)
Public FlagCtr As Integer ' ������� ������
Public Sub GetSosedi(Tek As Long, x As Long, Maximum As Long) '��������� � 1
' GetSosedi(�����_������_�_�������_Mines_���_�������_�����_������, ���-��_������_��_�����������, ���-��_������_��_����)
Dim I As Byte
For I = 0 To 8 '�������� ������ � ��������, � ���� ����� ��������!
    Sosedi(I) = 0
Next I
'��������� �� ������������� �������, ���� ��� �� ��� (������ � �����-���� ������� ����)
'���������� � ����� ���� ������ -1
'-----------------------------------------------
If (Tek / x) = (Int(Tek / x)) Then Sosedi(1) = -1: Sosedi(2) = -1: Sosedi(3) = -1
If ((Tek - 1) / x) = (Int((Tek - 1) / x)) Then Sosedi(5) = -1: Sosedi(6) = -1: Sosedi(7) = -1
If (Tek - x) <= 0 Then Sosedi(1) = -1: Sosedi(7) = -1: Sosedi(8) = -1
If (x + Tek) > Maximum Then Sosedi(3) = -1: Sosedi(4) = -1: Sosedi(5) = -1
'���������� ������ ������� (� ������� Mines)
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
'����� ������� ��� �������� ������ I - ������ ��� ������� ���������� �������

'                  H
'              G 7 8 1 (A)
'              F 6 I 2 (B)
'                5 4 3 (C)
'                E D

End Sub
Sub Inc(ByRef Include As Long, Couter As Long)
' ���������� � �������� Include �������� Couter ���� ��� <> 0, � ����� ���������� 1
If Couter = 0 Then Include = Include + 1 Else Include = Include + Couter
End Sub
Public Sub GameStart()
'������ �������� ����
Form1.Timer1.Enabled = True '�������� ������
Form1.Timer1.Interval = 1000 ' ������������� �������� ������������ � 1 ���.
End Sub
Public Sub SetFlags(flag As Integer)
' ��������� �������
' ���� ���� ��� ��� ���������� �� ������ ������: ������� ��� � ������� ������, ����������� �������� ���� "�������" ����������� (������� ������ ������), �������� ������� ������, ���������� ���������� ��� ���������������
If Flags(flag) Then Flags(flag) = False: Form1.Check1(flag).Picture = InvisItems.None.Picture: FlagCtr = FlagCtr - 1: Form1.Num1.Value = Form1.KolMines - FlagCtr: Exit Sub
' ������ ������ ��� ��� - ���������, �������� ��� ���
If FlagCtr >= Form1.KolMines Then CheckPobeda: Exit Sub
Flags(flag) = True '���������� ���� � ������� ������
Form1.Check1(flag).Picture = InvisItems.flag.Picture ' �������� �������� �� ���� �� ����������� ������
FlagCtr = FlagCtr + 1 ' ��������� ������� ������
Form1.Num1.Value = Form1.KolMines - FlagCtr ' ���������� ���������� ���
If FlagCtr = Form1.KolMines Then CheckPobeda: Exit Sub ' ���������, �������� ��� ���

End Sub
Sub CheckPobeda() ' ��������, ������� ���� ��� ������ ���������
Dim I As Integer
For I = 1 To Form1.MaxFiled '���� �� ���������� ������ �� ����
    '���� � ������� - ���� � ����� ��� ��� ���, ���� �� ����� 10
    If (Mines(I) = -1) And (Flags(I) = False) Then GoTo 10
Next I
Pobeda ' ����� ������!
Exit Sub
10 NoMines.Show 1 ' ���������� ������������ ��� �� ��� �� �����!
End Sub
Public Sub Porajen(Index As Integer)
Dim I As Integer
Form1.mnuGPorajen.Enabled = False '��������� ����������� ������� =)
For I = 1 To Form1.MaxFiled
'����, �� ������� �� �����������. ��� �������� ���� � checked � �������� - � ����
If I = Index Then Form1.Check1(I).DisabledPicture = InvisItems.mineboom.Picture: Form1.Check1(I).Value = Checked: Form1.Check1(I).Enabled = False: GoTo 10
'���� ��� ����� ���� - ������ �������� Check
If (Mines(I) = -1) And (Flags(I) = True) Then Form1.Check1(I).DisabledPicture = InvisItems.flag.Picture: Form1.Check1(I).Enabled = False: GoTo 10
'��� ����� ����� ���, ������ ���� � �������� Check
If (Mines(I) = -1) And (Flags(I) = False) Then Form1.Check1(I).DisabledPicture = InvisItems.Mine.Picture: Form1.Check1(I).Value = Checked: Form1.Check1(I).Enabled = False: GoTo 10
'���� ���� � ���� ���! �������!!! ������ ������� =)
If (Mines(I) <> -1) And (Flags(I) = True) Then Form1.Check1(I).DisabledPicture = InvisItems.errors.Picture: Form1.Check1(I).Value = Checked: Form1.Check1(I).Enabled = False: GoTo 10
Form1.Check1(I).Enabled = False ' �������� ������
10 Next I
Form1.Command1.Picture = InvisItems.Porajen.Picture ' �������� �� ������ ��. ����� ������ �� ������� ����
Form1.IsKonec = True '������������� ���� ����� ����
Form1.Timer1.Enabled = False ' ��������� ������
OpenDoor ' ��������� ����� CD (������ ������� ��������� ����� =)
If FileExist("boom.wav") Then PlayWAVE "boom.wav" ' ����������� ����� ���� ���� ����.
CloseDoor '��������� CD (������������ ������� ������ ������)
End Sub
Public Sub OpenNull(Index As Integer)
'��������� ��������� ������
Dim I As Integer '�������
Dim IsZero As Boolean '���� ������ ������
Dim Zero(750) As Long '���� ���������� ������ ������ ������� ����� � ���� ���������
Dim SPointer As Integer ' ��������� ��������� � ������� Zero
Dim TPointer As Integer ' ������� ��������� ��� ��
TPointer = 1 '������� �������������� 1
SPointer = 0 '��������� ��������� - 0
Form1.Check1(Index).Value = Checked ' ������������� ��� Index - �� �� ��� ������ � checked
Form1.Check1(Index).Enabled = False ' ��������� ���
yyy: SPointer = SPointer + 1 ' ���������� ��������� ��������� (� ������� Zero)
GetSosedi CLng(Index), Form1.UserX, Form1.MaxFiled '�������� ������� Index �� ������� ��� ������������� � Long - CLng
For I = 1 To 8
    If Sosedi(I) = -1 Then GoTo 10 ' ��� �������� ������ - ����� ��������� �������������
    If Flags(Sosedi(I)) Then GoTo 10 ' � �������� ������ ���� - ����������
    If Mines(Sosedi(I)) > 0 Then Cifra (Sosedi(I)): GoTo 10 ' � �������� ������ �����, ������ �� � ����������
    If Mines(Sosedi(I)) = -1 Then GoTo 10 ' � �������� ������ ���� - ����������
    If Mines(Sosedi(I)) = -2 Then GoTo 10 ' �������� ������ ���������������� - ����������
    Zero(TPointer) = Sosedi(I) ' ���������� ����� ������ ������ � ������
    Mines(Sosedi(I)) = -2 ' ������ �������, �� ��� ��� ���������������� ��� ������
    TPointer = TPointer + 1 ' ������� ��������� �����������
10 Next I
'���� � ������� Zero �� ���������� ��������� �� 0 - ������, ��
'����� �� ��� ������ ������� ����� �������. ����������� Index ��������� ������, ������� ����� �������������.
' � ��������� ��������� �������, ������ �� yyy
If Zero(SPointer) <> 0 Then Index = Zero(SPointer): GoTo yyy
' ��������� ���� �� �����
For I = 1 To 750
    If Zero(I) = 0 Then GoTo 20 '����� 0 ��-�  ������� Zero - ������ ��������� ����, ������
    Form1.Check1(Zero(I)).Value = Checked ' ������������� ��� �� ���������� ���������
    Form1.Check1(Zero(I)).Enabled = False ' ��������� ���
Next I
20
End Sub
Sub Cifra(Index As Integer)
'����������� ����� �� ����� ���� �� �����
Form1.Check1(Index).DisabledPicture = InvisItems.Numers(Mines(Index)).Picture
' ������������� � �������� ���� (DisabledPicture - �������� � ����������� ���������) � ������������
' � ������ ��� ����� ���� (�� ������� Mines) �������� �� ������� ��������� Pic4ureBox Numers
Form1.Check1(Index).Value = Checked ' ������������� ��� �� ���������� ���������
Form1.Check1(Index).Enabled = False ' ��������� ���
End Sub
Sub Pobeda()
' �������� ��� ������
Dim I As Integer '�������
For I = 1 To Form1.MaxFiled '������ ����� � ��������� ����
    ' ���� ���� ���� - ������ ��� ��� ���� - ������������� ������� �����, �������� � ��������� ���
    If Mines(I) = -1 Then Flags(I) = True: Form1.Check1(I).Picture = InvisItems.flag.Picture: _
    Form1.Check1(I).DisabledPicture = InvisItems.flag.Picture
    Form1.Check1(I).Enabled = False
Next I
Form1.IsKonec = True ' ������� ����� ����
Form1.Command1.Picture = InvisItems.Pobeda.Picture ' �� �������� �� ����� ������ �������� ������� � �����
Form1.Timer1.Enabled = False ' ��������� ������
If FileExist("pobeda.wav") Then PlayWAVE "pobeda.wav" ' ���� ���� �������� ���� - �����������
End Sub
