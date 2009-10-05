Dim pdfexecVersion, needCScript
pdfexecVersion = "1.5"
needCScript = False

Sub ShowError(msg)
    DoLog "ERROR:   " & replace(replace(msg, vbCrLf & vbCrLf, vbCrLf), vbCrLf, vbCrLf & "         ")
    If Not noDialogs Then
        MsgBox msg, 16, "PDFexec v" & pdfexecVersion & ": Error"
    ElseIf Not noBeep Then
        ' Beep!
        Wscript.StdOut.Write chr(007)
    End If
End Sub

Sub ShowWarning(msg)
    DoLog "WARNING: " & replace(replace(msg, vbCrLf & vbCrLf, vbCrLf), vbCrLf, vbCrLf & "         ")
    If Not noDialogs Then
        MsgBox msg, 48, "PDFexec v" & pdfexecVersion & ": Warning"
    End If
End Sub

Sub DoProgress()
    If hasConsole Then
        Wscript.StdOut.Write "."
    End If
End Sub

Sub DoProgressDone()
    If hasConsole Then
        Wscript.StdOut.WriteLine ""
    End If
End Sub

Sub DoLog(msg)
    If hasConsole Then
        Wscript.StdOut.WriteLine msg
    End If
End Sub

Sub ShowMessage(msg)
    DoLog replace(msg, vbCrLf & vbCrLf, vbCrLf)
    If Not noDialogs Then
        MsgBox msg, 64, "PDFexec v" & pdfexecVersion
    End If
End Sub

Function CheckStartMode()
    ' Returns the running executable as upper case from the last \ symbol
    strStartExe = UCase( Mid( wscript.fullname, instrRev(wscript.fullname, "\") + 1 ) )
    If strStartExe = "CSCRIPT.EXE" Then
        hasConsole = true
    ElseIf needCScript Then
        CheckStartMode = True
    End If
End Function

Sub RestartWithCScript()
    ' wscript.scriptfullname is the full path to the actual script
    Set oSh = CreateObject("wscript.shell")
    oSh.Run "cscript.exe /nologo """ & wscript.scriptfullname & """ " & args
    wscript.quit
End Sub

Set fso = CreateObject("Scripting.FileSystemObject")
Set exec = CreateObject("WScript.Shell")

Dim hasConsole
hasConsole = false
' Check that the starup mode is OK
If CheckStartMode() Then
    RestartWithCScript
End If

'Dim args As String
Dim args, file, pause, pauseErrors, noDialogs, noLogo
noLogo = False
pause = False
pauseErrors = False
noDialogs = False
noBeep = False
Set argList = WScript.Arguments
For Each arg In argList
    Select Case arg
    Case "\nologo"
        noLogo = True
    Case "\nobeep"
        noBeep = True
    Case "\s"
        noDialogs = True
    Case "\p"
        pause = True
    Case "/p"
        pause = True
    Case "\pe"
        pauseErrors = True
    Case "/pe"
        pauseErrors = True
    Case Else
        If Left(arg, 1) = "\" Or Left(arg, 1) = "/" Then
            ShowWarning "Unknown argument: " & arg
        Else
            file = arg
        End If
    End Select
    args = args & arg & " "
Next

If Not noLogo Then
    DoLog "==============================================================================="
    DoLog "|              | Intelligent wrapper for LaTeX.                               |"
    DoLog "| PDFexec v" & pdfexecVersion & " | Copyright (c) 2009, Paul Lammertsma                          |"
    DoLog "|              | Freely available from http://paul.luminos.nl/                |"
    DoLog "|-----------------------------------------------------------------------------|"
    DoLog "| For usage and command line switches, start PDFexec with the /? switch.      |"
    DoLog "===============================================================================" & vbCrLf
End If

If hasConsole Then
    DoLog "Using CScript console."
End If

If file = "" Then
    ShowError "No input file was specified."
    wscript.quit 2
End If

texFile = file & ".tex"
pdfFile = file & ".pdf"

If fso.FileExists(pdfFile) Then
    exec.Exec "pdfclose.exe --file=" & pdfFile
Else
    exec.Exec "pdfclose.exe"
End If
exec.Exec "taskkill /FI ""WINDOWTITLE eq " & pdfFile & " - Adobe Acrobat Professional"""
exec.Exec "taskkill /FI ""WINDOWTITLE eq " & pdfFile & " - Adobe Acrobat"""
exec.Exec "taskkill /FI ""WINDOWTITLE eq " & pdfFile & " - Adobe Reader"""

hadError = False

If fso.FileExists(texFile) Then
    DoLog "Compiling LaTeX file: " & texFile
    'Set response = exec.Run("pdflatex -interaction=nonstopmode -c-style-errors """ & texFile & """", 1, True)
    Set response = exec.Exec("pdflatex -interaction=nonstopmode -halt-on-error """ & texFile & """")
    errors = ""
    finalError = ""
    emergency = False
    Do While Not response.StdOut.AtEndOfStream
        'strPResult=replace(oexec.StdOut.ReadLine, vbLF, "<br /><hr />"
        'response.flush
        errLine = response.StdOut.ReadLine()
        If Left(errLine, 6) = "!  ==>" Then
            finalError = vbCrLf & vbCrLf & Mid(errLine, 8)
        ElseIf errLine = "! Emergency stop." Then
            emergency = True
        ElseIf Left(errLine, 15) = "! LaTeX Error: " Then
            errors = errors & vbCrLf & "  - " & Mid(errLine, 16)
        ElseIf Left(errLine, 1) = "!" Then
            errors = errors & vbCrLf & "  - " & Mid(errLine, 3)
            DoLog errLine
        Else
            'DoLog errLine
            DoProgress
        End If
    Loop
    DoProgressDone
    If emergency Or finalError <> "" Then
        hadError = True
        If errors = "" And finalError = "" Then
            ShowError "An unspecified error occured during compilation. (Error #" & response.ExitCode & ")"
        ElseIf errors = "" Then
            ShowError "An error occurred during compilation:" & finalError & " (Error #" & response.ExitCode & ")"
        Else
            ShowError "LaTeX returned the following error(s) during compilation:" & errors & finalError & " (Error #" & response.ExitCode & ")"
        End If
    Else
        If errors = "" And finalError <> "" Then
            ShowWarning "An error occurred during compilation:" & finalError & " (Error #" & response.ExitCode & ")"
        ElseIf errors <> "" Then
            ShowWarning "LaTeX returned the following error(s) during compilation:" & errors & finalError & " (Error #" & response.ExitCode & ")"
        End If
    End If
    If Not hadError Then
        exec.Exec "pdfopen.exe --file=" & pdfFile
    End If
Else
    hadError = True
    ShowError "Input file was missing:" & vbCrLf & vbCrLf & texFile
End If

If pause Or (hadError And pauseErrors) Then
    Wscript.StdOut.Write "Press return..."
    Do While Not WScript.StdIn.AtEndOfLine
       Input = WScript.StdIn.Read(1)
    Loop
End If

wscript.quit
