Attribute VB_Name = "winmm"
Option Explicit 'контроль объ€влени€ переменных
'объ€вл€ем API функции. ќписание декларации вз€то из программы
' API Text Viewer в комплекте VB
' импортируем функцию воспроизведени€ wave-файлов 1 параметр, им€ файла 2-й флаги (не понадоб€ццо)
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
' mciSendString посылка управл€ющей команды драйверам мультимедийных устройств - CD-ROM'у, MIDI-секвенсеру
' 1 параметр - команда
' 2 - адрес буфера ответа устройства
' 3 - длина буфера дл€ ответа
' указатель на вызвыающее приложение =)
Public Declare Function mciSendString Lib "winmm.dll" _
                  Alias "mciSendStringA" _
                  (ByVal lpstrCommand As String, _
                  ByVal lpstrReturnString As String, _
                  ByVal uReturnLength As Long, _
                  ByVal hwndCallback As Long) As Long
'ќткрыть дверцу CD-Rom:
'Call mciSendString("Set CDAudio Door Open Wait", 0&, 0&, 0&)
'«акрыть дверцу CD-Rom:
'Call mciSendString("Set CDAudio Door Closed Wait", 0&, 0&, 0&)

Public Sub PlayWAVE(FileName As String)
' функци€ проигрывани€ wave файла.
sndPlaySound FileName, 0
End Sub
Public Sub OpenDoor()
' открываем дверь CD
Call mciSendString("Set CDAudio Door Open Wait", 0&, 0&, 0&)
End Sub
Public Sub CloseDoor()
' закрываем дверь CD
Call mciSendString("Set CDAudio Door Closed Wait", 0&, 0&, 0&)
End Sub
