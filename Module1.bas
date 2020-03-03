Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Sub main()
    Do
        Clipboard.Clear
        Wait 100
    Loop
End Sub
'延时，单位为毫秒
Public Function Wait(ByVal MilliSeconds As Long)
    Dim dSavetime As Double
    dSavetime = timeGetTime + MilliSeconds   '记下开始时的时间
    While timeGetTime < dSavetime '循环等待
        DoEvents '转让控制权，以便让操作系统处理其它的事件
    Wend
End Function
