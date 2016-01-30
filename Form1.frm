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
   StartUpPosition =   3  'Windows �⺻��
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



Private Sub Form_Load()
Dim a As TIME_ZONE_INFORMATION
Dim b As Integer
b = GetTimeZoneInformation(a)
'MsgBox a.Bias / 60 'GMT + ������ Ȯ��
'MsgBox a.StandardBias '�ǹ̾���
'MsgBox a.StandardName(31) '�ǹ̾���
'MsgBox a.DaylightBias / 60 '�ϱ��ð� ���� ����// �����Ǿ������� 1

Select Case a.Bias / 60
Case 6
SaveSetting "TimeZoneChanger", "DaylightBias", "Settings", a.DaylightBias '�ϱ����� ���� ������ �����ϱ�
a.Bias = -9 * 60
a.DaylightBias = 0 '���ѹα��� �ϱ����� �������
SetTimeZoneInformation a
MsgBox "���ѹα� ǥ�ؽ÷� ���� �Ϸ�", vbInformation, "TimeZone"

Case -9
x = 0
a.DaylightBias = GetSetting("TimeZoneChanger", "DaylightBias", "Settings") '�ϱ����� ���� ������ �޾ƿ���

a.Bias = 6 * 60
SetTimeZoneInformation a
DeleteSetting "Timezonechanger"
MsgBox "�̱� �ߺ� ǥ�ؽ÷� ���� �Ϸ�", vbInformation, "TimeZone"
End Select
End
End Sub
