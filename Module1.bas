Attribute VB_Name = "Module1"
'===============================================================
'                 WIN32 API함수 및 구조체 선언부
'===============================================================
Private Declare Function SetTimeZoneInformation Lib "kernel32" _
(lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Function GetTimeZoneInformation Lib "kernel32" _
(lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Private Type SYSTEMTIME '시스템 시간 구조체
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION '타임존 구조체
    Bias As Long 'GMT + Bias
    StandardName(31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long 'Daylight Time 여부
End Type
'======================== 선언부 끝 ========================


Sub Main()
Dim a As TIME_ZONE_INFORMATION 'Timezoneinformation 구조체를 변수로 선언
Dim b As Integer 'Timezoneinformation의 반환값
b = GetTimeZoneInformation(a)

Select Case a.Bias / 60
Case 6
SaveSetting "TimeZoneChanger", "DaylightBias", "Settings", a.DaylightBias '일광절약 여부 데이터 저장하기
a.Bias = -9 * 60
a.DaylightBias = 0 '대한민국은 일광절약 적용안함
SetTimeZoneInformation a

If App.PrevInstance = True Then
Shell "taskkill /im ChangeTime.exe /t /f", vbHide
End If
MsgBox "대한민국 표준시로 변경 완료", vbInformation, "TimeZone"

Case -9
x = 0
a.DaylightBias = GetSetting("TimeZoneChanger", "DaylightBias", "Settings") '일광절약 여부 데이터 받아오기

a.Bias = 6 * 60
SetTimeZoneInformation a
DeleteSetting "Timezonechanger"

If App.PrevInstance = True Then '두번 실행될시 두 프로그램 다 종료
Shell "taskkill /im ChangeTime.exe /t /f", vbHide
End If
MsgBox "미국 중부 표준시로 변경 완료", vbInformation, "TimeZone"
End Select
End
End Sub
