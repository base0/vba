' code from ChatGPT
Sub RecalculateEverySecond()
    Application.OnTime Now + TimeValue("00:00:01"), "RecalculateEverySecond"
    Application.Calculate
End Sub
