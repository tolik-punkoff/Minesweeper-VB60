Attribute VB_Name = "systemfunction"
Option Explicit '�������� ���������� ����������

Public Function FileExist(FileEx$) As Boolean
' �������� ������� �����
Dim ff As Long ' ���������� ��� ��������� �������� �����
On Error GoTo 10 ' ���� ������ - �� ����� ���, ���� �����
ff = FreeFile() ' ������� ��������� �������� �����
Open FileEx$ For Input As ff ' �������� ������� ���� ��� ������
Close ff ' ���������
FileExist = True ' ���� ������ �� ���� �� ���������, ���� ���� - ���. �� �-�� � TRUE
Exit Function ' ������� �� �������
10 FileExist = False ' ����� ������������� �������� � FALSE
End Function

