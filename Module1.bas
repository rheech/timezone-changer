Attribute VB_Name = "Module1"
'===============================================================
'                 WIN32 API�Լ� �� ����ü �����
'===============================================================
Private Declare Function SetTimeZoneInformation Lib "kernel32" _
(lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Function GetTimeZoneInformation Lib "kernel32" _
(lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Private Type SYSTEMTIME '�ý��� �ð� ����ü
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION 'Ÿ���� ����ü
    Bias As Long 'GMT + Bias
    StandardName(31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long 'Daylight Time ����
End Type
'======================== ����� �� ========================


Sub Main()
Dim a As TIME_ZONE_INFORMATION 'Timezoneinformation ����ü�� ������ ����
Dim b As Integer 'Timezoneinformation�� ��ȯ��
b = GetTimeZoneInformation(a)

Select Case a.Bias / 60
Case 6
SaveSetting "TimeZoneChanger", "DaylightBias", "Settings", a.DaylightBias '�ϱ����� ���� ������ �����ϱ�
a.Bias = -9 * 60
a.DaylightBias = 0 '���ѹα��� �ϱ����� �������
SetTimeZoneInformation a

If App.PrevInstance = True Then
Shell "taskkill /im ChangeTime.exe /t /f", vbHide
End If
MsgBox "���ѹα� ǥ�ؽ÷� ���� �Ϸ�", vbInformation, "TimeZone"

Case -9
x = 0
a.DaylightBias = GetSetting("TimeZoneChanger", "DaylightBias", "Settings") '�ϱ����� ���� ������ �޾ƿ���

a.Bias = 6 * 60
SetTimeZoneInformation a
DeleteSetting "Timezonechanger"

If App.PrevInstance = True Then '�ι� ����ɽ� �� ���α׷� �� ����
Shell "taskkill /im ChangeTime.exe /t /f", vbHide
End If
MsgBox "�̱� �ߺ� ǥ�ؽ÷� ���� �Ϸ�", vbInformation, "TimeZone"
End Select
End
End Sub
