Attribute VB_Name = "winmm"
Option Explicit '�������� ���������� ����������
'��������� API �������. �������� ���������� ����� �� ���������
' API Text Viewer � ��������� VB
' ����������� ������� ��������������� wave-������ 1 ��������, ��� ����� 2-� ����� (�� �����������)
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
' mciSendString ������� ����������� ������� ��������� �������������� ��������� - CD-ROM'�, MIDI-����������
' 1 �������� - �������
' 2 - ����� ������ ������ ����������
' 3 - ����� ������ ��� ������
' ��������� �� ���������� ���������� =)
Public Declare Function mciSendString Lib "winmm.dll" _
                  Alias "mciSendStringA" _
                  (ByVal lpstrCommand As String, _
                  ByVal lpstrReturnString As String, _
                  ByVal uReturnLength As Long, _
                  ByVal hwndCallback As Long) As Long
'������� ������ CD-Rom:
'Call mciSendString("Set CDAudio Door Open Wait", 0&, 0&, 0&)
'������� ������ CD-Rom:
'Call mciSendString("Set CDAudio Door Closed Wait", 0&, 0&, 0&)

Public Sub PlayWAVE(FileName As String)
' ������� ������������ wave �����.
sndPlaySound FileName, 0
End Sub
Public Sub OpenDoor()
' ��������� ����� CD
Call mciSendString("Set CDAudio Door Open Wait", 0&, 0&, 0&)
End Sub
Public Sub CloseDoor()
' ��������� ����� CD
Call mciSendString("Set CDAudio Door Closed Wait", 0&, 0&, 0&)
End Sub
