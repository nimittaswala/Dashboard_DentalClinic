Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

Dim pyPath, localApp
pyPath = ""
localApp = WshShell.ExpandEnvironmentStrings("%LOCALAPPDATA%")

Dim candidates(15)
candidates(0)  = localApp & "\Programs\Python\Python314\python.exe"
candidates(1)  = localApp & "\Programs\Python\Python313\python.exe"
candidates(2)  = localApp & "\Programs\Python\Python312\python.exe"
candidates(3)  = localApp & "\Programs\Python\Python311\python.exe"
candidates(4)  = "C:\Program Files\Python314\python.exe"
candidates(5)  = "C:\Program Files\Python313\python.exe"
candidates(6)  = "C:\Program Files\Python312\python.exe"
candidates(7)  = "C:\Program Files\Python311\python.exe"
candidates(8)  = "C:\Program Files (x86)\Python314\python.exe"
candidates(9)  = "C:\Program Files (x86)\Python313\python.exe"
candidates(10) = "C:\Program Files (x86)\Python312\python.exe"
candidates(11) = "C:\Program Files (x86)\Python311\python.exe"
candidates(12) = "C:\Python314\python.exe"
candidates(13) = "C:\Python313\python.exe"
candidates(14) = "C:\Python312\python.exe"
candidates(15) = "C:\Python311\python.exe"

Dim i
For i = 0 To 15
    If fso.FileExists(candidates(i)) Then
        pyPath = candidates(i)
        Exit For
    End If
Next

If pyPath = "" Then
    MsgBox "Python not found. Please run install.bat first.", 16, "Dental Dashboard"
    WScript.Quit 1
End If

WshShell.Run Chr(34) & pyPath & Chr(34) & " " & Chr(34) & "C:\dental_automation\dashboard\dashboard.py" & Chr(34), 0, False
WshShell.Run Chr(34) & pyPath & Chr(34) & " " & Chr(34) & "C:\dental_automation\ar\ar_dashboard.py" & Chr(34), 0, False
Set WshShell = Nothing
