VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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



Private Sub Form_Load()
Dim a As TIME_ZONE_INFORMATION
Dim b As Integer
b = GetTimeZoneInformation(a)
'MsgBox a.Bias / 60 'GMT + 몇인지 확인
'MsgBox a.StandardBias '의미없음
'MsgBox a.StandardName(31) '의미없음
'MsgBox a.DaylightBias / 60 '일광시간 절약 여부// 설정되어있으면 1

Select Case a.Bias / 60
Case 6
SaveSetting "TimeZoneChanger", "DaylightBias", "Settings", a.DaylightBias '일광절약 여부 데이터 저장하기
a.Bias = -9 * 60
a.DaylightBias = 0 '대한민국은 일광절약 적용안함
SetTimeZoneInformation a
MsgBox "대한민국 표준시로 변경 완료", vbInformation, "TimeZone"

Case -9
x = 0
a.DaylightBias = GetSetting("TimeZoneChanger", "DaylightBias", "Settings") '일광절약 여부 데이터 받아오기

a.Bias = 6 * 60
SetTimeZoneInformation a
DeleteSetting "Timezonechanger"
MsgBox "미국 중부 표준시로 변경 완료", vbInformation, "TimeZone"
End Select
End
End Sub
