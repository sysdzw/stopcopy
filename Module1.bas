Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Sub main()
    Do
        Clipboard.Clear
        Wait 100
    Loop
End Sub
'��ʱ����λΪ����
Public Function Wait(ByVal MilliSeconds As Long)
    Dim dSavetime As Double
    dSavetime = timeGetTime + MilliSeconds   '���¿�ʼʱ��ʱ��
    While timeGetTime < dSavetime 'ѭ���ȴ�
        DoEvents 'ת�ÿ���Ȩ���Ա��ò���ϵͳ�����������¼�
    Wend
End Function
