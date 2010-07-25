Dim pdfexecVersion, needCScript
pdfexecVersion = "1.6"
needCScript = False
needRerun = "Label(s) may have changed. Rerun"
needsRerun = False
autoRerun = True
pdfFile = ""
run = 1

Set oShell = CreateObject("WScript.Shell")

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

Sub ShowMessage(msg)
    DoLog replace(msg, vbCrLf & vbCrLf, vbCrLf)
    If Not noDialogs Then
        MsgBox msg, 64, "PDFexec v" & pdfexecVersion
    End If
End Sub

Sub DoProgressStart(msg)
    If hasConsole And msg <> "" Then
        Wscript.StdOut.Write msg
    End If
End Sub

Sub DoProgress()
    If hasConsole Then
        Wscript.StdOut.Write "."
    End If
End Sub

Sub DoProgressDone()
    If hasConsole Then
        Wscript.StdOut.WriteLine " Done."
    End If
End Sub

Sub DoLog(msg)
    If hasConsole Then
        Wscript.StdOut.WriteLine msg
    End If
End Sub

Function NoPeriod(msg)
    If Right(msg, 1) = "." Or Right(msg, 1) = "!" Then
        msg = Mid(msg, 1, Len(msg)-1)
    End If
    NoPeriod = msg
End Function

Sub OpenPDF(file)
    DoProgressStart "Opening " & file & ".."
    openErr = ""
    Set response = oShell.Exec("pdfopen.exe --file=" & file)
    Do While response.Status = 0
    'Do While Not response.StdErr.AtEndOfStream
        DoProgress
        If response.Status > 0 Then
            Exit Do
        End If
        WScript.Sleep 250
        'openErr = openErr & vbCrLf & response.StdErr.ReadLine()
    Loop
    If response.ExitCode = 0 Then
        DoProgressDone
    Else
        DoLog " Failed."
        hasError = True
        ShowError "Failed opening """ & file & """ from directory:" & vbCrLf & vbCrLf & "   " & oShell.CurrentDirectory
    End If
End Sub

Sub ClosePDF(file)
    DoProgressStart "Closing PDF.."
    If fso.FileExists(pdfFile) Then
        Set response1 = oShell.Exec("pdfclose.exe --file=" & file)
    Else
        Set response1 = oShell.Exec("pdfclose.exe")
    End If
    Set response2 = oShell.Exec("taskkill /FI ""WINDOWTITLE eq " & file & " - Adobe Acrobat Professional""")
    Set response3 = oShell.Exec("taskkill /FI ""WINDOWTITLE eq " & file & " - Adobe Acrobat""")
    Set response4 = oShell.Exec("taskkill /FI ""WINDOWTITLE eq " & file & " - Adobe Reader""")
    
    Do While response1.Status = 0 Or response2.Status = 0 Or response3.Status = 0 Or response4.Status = 0
        DoProgress
        WScript.Sleep 250
    Loop
    DoProgressDone
    oShell.AppActivate "PDFExec"
End Sub

Function HandleError(output, filename, pos)
    mode = 0
    msg = ""
    If Left(output, 6) = "!  ==>" Then
        msg = NoPeriod(Mid(output, 8))
    ElseIf Left(output, 7) = "!   ==>" Then
        msg = NoPeriod(Mid(output, 9))
    ElseIf output = "! Emergency stop." Then
        emergency = True
        Exit Function
    ElseIf Left(output, 15) = "! LaTeX Error: " Then
        mode = 2
        msg = NoPeriod(Mid(output, 16))
    ElseIf Left(output, 9) = "Warning: " Then
        mode = 1
        msg = NoPeriod(Mid(output, 10))
    ElseIf Left(output, 15) = "LaTeX Warning: " Then
        mode = 1
        msg = NoPeriod(Mid(output, 16))
    ElseIf Left(output, 17) = "! LaTeX Warning: " Then
        mode = 1
        msg = NoPeriod(Mid(output, 18))
    ElseIf Left(output, 1) = "!" Then
        mode = 2
        msg = NoPeriod(Mid(output, 3))
    Else
        Exit Function
    End If
    If mode > 0 Then
        If pos > 0 Then
            msg = msg & " on input line " & pos
        End If
        If filename <> "" Then
            msg = msg & " in """ & filename & """"
        End If
    End If
    If mode = 1 Then
        If Left(msg, Len(needRerun)) = needRerun Then
            needsRerun = True
        Else
            warnings = warnings & vbCrLf & "  - " & msg
            warningCount = warningCount + 1
        End If
    ElseIf mode = 2 Then
        errors = errors & vbCrLf & "  - " & msg
        errorCount = errorCount + 1
    Else
        fatalError = fatalError & vbCrLf & vbCrLf & msg & "."
        If pos > 0 Then
            fatalErrorLine = pos
        End If
    End If
End Function

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
    oShell.Run "cscript.exe /nologo """ & wscript.scriptfullname & """ " & args
    wscript.quit
End Sub

Sub pdfExec()
    DoLog "Path:         " & pathFile
    If pathFile <> "" And Not fso.FolderExists(pathFile) Then
        hadError = True
        ShowError "The path cannot be found:" & vbCrLf & vbCrLf & pathFile
        Exit Sub
    End If
    needsRerun = False
    If fso.FileExists(texFile) Then
        If run = 1 Then
            Set objFile = fso.GetFile(texFile)
            If objFile.Size > 0 Then
                Set objReadFile = fso.OpenTextFile(texFile, 1)
                Do While Not objReadFile.AtEndOfStream
                    fileLine = objReadFile.ReadLine()
                    If Left(fileLine, 1) = "%" Then
                        fileLine = Trim(Mid(fileLine, 2))
                        If Left(fileLine, 7) = "parent=" Then
                            fileLine = Mid(fileLine, 8)
                            parentSteps = parentSteps + 1
                            If parentSteps > 4 Then
                                hadError = True
                                ShowError "More than 4 parent files were referenced. Check your file headers for recursion."
                                Exit Sub
                            End If
                            DoLog "Parent file:  " & fileLine
                            fileLine = Replace(fileLine, "/", "\")
                            pos = InStrRev(fileLine, "\")
                            If pos > 0 Then
                                pathFile = fso.GetAbsolutePathName(pathFile & Mid(fileLine, 1, pos))
                                endChr = Right(pathFile, 1)
                                If endChr <> "/" And endChr <> "\" Then pathFile = pathFile & "\"
                                oShell.CurrentDirectory = pathFile
                                texFile = Mid(fileLine, pos+1)
                            Else
                                texFile = fileLine
                            End If
                            pdfExec()
                            Exit Sub
                        End If
                    Else
                        Exit Do
                    End If
                Loop
            Else
                hadError = True
                ShowError "Input file was empty:" & vbCrLf & vbCrLf & texFile
                Exit Sub
            End If
            pos = InStrRev(texFile, ".")
            If pos > 0 Then
                pdfFile = Left(texFile, pos-1) & ".pdf"
                filename = Left(texFile, pos-1)
            Else
                pdfFile = texFile & ".pdf"
                filename = texFile
            End If
            bibFile = filename & ".bib"
            ClosePDF pdfFile
        End If
        
        DoLog ""
        DoLog "Run #" & run & "..."

        errors = ""
        errorCount = 0
        warnings = ""
        warningCount = 0
        fatalError = ""
        fatalErrorLine = 0
        fullOutput = ""
        emergency = False
        errPrefix = pathFile
        errContinues = False

        If fso.FileExists(bibFile) Then
            DoProgressStart "Compiling BibTeX file: " & bibFile
            If verbose Then
                DoLog "..."
            End If
            Set response = oShell.Exec("bibtex """ & filename & """")
            Do While Not response.StdOut.AtEndOfStream
                If response.Status > 0 Then
                    Exit Do
                End If
                curLine = response.StdOut.ReadLine()
                ' fullOutput = fullOutput & vbCrLf & errLine
                If errContinues Then
                    errLine = errLine & curLine
                    If Len(curLine) < 79 Then
                        errContinues = False
                    End If
                Else
                    errLine = curLine
                    If Len(curLine) >= 79 Then
                        errContinues = True
                    End If
                End If
                If Not errContinues Then
                    If verbose Then
                        DoLog errLine
                    Else
                        DoProgress
                    End If
                    HandleError errLine, bibFile, 0
                    If Left(errLine, Len(errPrefix)) = errPrefix Then
                        errLine = Mid(errLine, Len(errPrefix)+1)
                        pos1 = InStr(errLine, ":")
                        pos2 = InStr(errLine, " ")
                        errFile = ""
                        If (pos1 > 0 And pos1 < 6) Or (pos1 > 0 And pos2 > 0 And pos1 < pos2) Then
                            errFile = Left(errLine, pos1-1)
                            If pos1 > 0 And pos1<pos2 Then
                                errLine = Mid(errLine, pos1+1)
                            Else
                                errLine = Mid(errLine, pos2+1)
                            End If
                        End If
                        errPos = 0
                        pos1 = InStr(errLine, ":")
                        pos2 = InStr(errLine, " ")
                        If (pos1 > 0 And pos1 < 6) Or (pos1 > 0 And pos2 > 0 And pos1 < pos2) Then
                            errPos = Left(errLine, pos1-1)
                            If pos1 > 0 And pos1<pos2 Then
                                errLine = Mid(errLine, pos1+1)
                            Else
                                errLine = Mid(errLine, pos2+1)
                            End If
                        End If
                        HandleError "! " & errLine, errFile, CInt(errPos)
                    End If
                End If
            Loop
            If verbose Then
                DoLog ""
            Else
                DoProgressDone
            End If
            errCode = response.ExitCode
        End If

        If errCode = 0 And Not emergency And fatalError = "" Then
            DoProgressStart "Compiling LaTeX file: " & texFile
            If verbose Then
                DoLog "..."
            End If
            Set response = oShell.Exec("pdflatex -interaction=nonstopmode -c-style-errors -halt-on-error """ & texFile & """")
            Do While Not response.StdOut.AtEndOfStream
                If response.Status > 0 Then
                    Exit Do
                End If
                curLine = response.StdOut.ReadLine()
                ' fullOutput = fullOutput & vbCrLf & errLine
                If errContinues Then
                    errLine = errLine & curLine
                    If Len(curLine) < 79 Then
                        errContinues = False
                    End If
                Else
                    errLine = curLine
                    If Len(curLine) >= 79 Then
                        errContinues = True
                    End If
                End If
                If Not errContinues Then
                    If verbose Then
                        DoLog errLine
                    Else
                        DoProgress
                    End If
                    HandleError errLine, texFile, 0
                    If Left(errLine, Len(errPrefix)) = errPrefix Then
                        errLine = Mid(errLine, Len(errPrefix)+1)
                        pos1 = InStr(errLine, ":")
                        pos2 = InStr(errLine, " ")
                        errFile = ""
                        If (pos1 > 0 And pos1 < 6) Or (pos1 > 0 And pos2 > 0 And pos1 < pos2) Then
                            errFile = Left(errLine, pos1-1)
                            If pos1 > 0 And pos1<pos2 Then
                                errLine = Mid(errLine, pos1+1)
                            Else
                                errLine = Mid(errLine, pos2+1)
                            End If
                        End If
                        errPos = 0
                        pos1 = InStr(errLine, ":")
                        pos2 = InStr(errLine, " ")
                        If (pos1 > 0 And pos1 < 6) Or (pos1 > 0 And pos2 > 0 And pos1 < pos2) Then
                            errPos = Left(errLine, pos1-1)
                            If pos1 > 0 And pos1<pos2 Then
                                errLine = Mid(errLine, pos1+1)
                            Else
                                errLine = Mid(errLine, pos2+1)
                            End If
                        End If
                        HandleError "! " & errLine, errFile, CInt(errPos)
                    End If
                End If
            Loop
            If verbose Then
                DoLog ""
            Else
                DoProgressDone
            End If
            DoLog ""
        End If

        errCode = response.ExitCode
        If errCode <> 0 Or emergency Or fatalError <> "" Then
            hadError = True
            If errors = "" And fatalError = "" Then
                ShowError "An unspecified fatal error occured during compilation."
            ElseIf errors = "" Then
                If fatalErrorLine > 0 Then
                    ShowError "A fatal error occurred on input line " & fatalErrorLine & ":" & fatalError
                Else
                    ShowError "A fatal error occurred during compilation:" & fatalError
                End If
            Else
                msg = "LaTeX returned the following error(s) during compilation:"
                If fatalErrorLine > 0 Then
                    msg = "A fatal error occurred on input line " & fatalErrorLine & ":"
                End If
                ShowError msg & vbCrLf & errors & fatalError
            End If
        Else
            If errors = "" And fatalError <> "" Then
                hadError = True
                If fatalErrorLine > 0 Then
                    ShowError "A warning occurred on input line " & fatalErrorLine & ":" & fatalError
                Else
                    ShowError "A warning occurred during compilation:" & fatalError
                End If
            ElseIf errors <> "" Then
                warnings = errors & vbCrLf & warnings
            End If
        End If
        If needsRerun Then
            doRerun = False
            If autoRerun Then
                doRerun = True
            ElseIf MsgBox("A second pass is needed to complete compilation. Do you want to run it now?", 48 + 4, "PDFexec v" & pdfexecVersion) = vbYes Then
                doRerun = True
            End If
            If doRerun Then
                run = run + 1
                pdfExec()
                Exit Sub
            End If
        End If
        If warnings <> "" And showWarnings Then
            hadWarning = True
            ShowWarning "LaTeX returned the following warning(s) during compilation:" & vbCrLf & warnings & fatalError
            DoLog ""
        End If
        DoLog "Finished with " & errorCount & " errors and " & warningCount & " warnings."
        DoLog ""
        If Not hadError Then
            OpenPDF pdfFile
        End If
    Else
        hadError = True
        ShowError "Input file was missing:" & vbCrLf & vbCrLf & texFile
    End If
End Sub

Set fso = CreateObject("Scripting.FileSystemObject")

Dim hasConsole
hasConsole = false
' Check that the starup mode is OK
If CheckStartMode() Then
    RestartWithCScript
End If

'Dim args As String
Dim args, file, pause, pauseErrors, noDialogs, noLogo
noLogo = False
pauseSuccess = False
pauseWarnings = False
pauseErrors = False
showSuccess = False
showWarnings = False
showErrors = True
noDialogs = False
noBeep = False
verbose = False
sleep = 0
Set argList = WScript.Arguments
For Each arg In argList
    If Left(arg, 7) = "\pause:" Then
        mode = Mid(arg, 8)
        If IsNumeric(mode) Then
            mode = CInt(mode)
            If (mode And 1) > 0 Then
                pauseSuccess = True
            End If
            If (mode And 2) > 0 Then
                pauseWarnings = True
            End If
            If (mode And 4) > 0 Then
                pauseErrors = True
            End If
        End If
    ElseIf Left(arg, 6) = "\wait:" Then
        mode = Mid(arg, 7)
        If IsNumeric(mode) Then
            sleep = mode
        End If
    ElseIf Left(arg, 6) = "\show:" Then
        mode = Mid(arg, 7)
        If IsNumeric(mode) Then
            mode = CInt(mode)
            showWarnings = False
            showErrors = False
            If (mode And 1) > 0 Then
                showSuccess = True
            End If
            If (mode And 2) > 0 Then
                showWarnings = True
            End If
            If (mode And 4) > 0 Then
                showErrors = True
            End If
        End If
    Else
        Select Case arg
        Case "\nologo"
            noLogo = True
        Case "\nobeep"
            noBeep = True
        case "\norerun"
            autoRerun = False
        Case "\s"
            noDialogs = True
        Case "\v"
            verbose = True
        Case "\p"
            pauseSuccess = True
        Case "/p"
            pauseSuccess = True
        Case "/pe"
            pauseErrors = True
            pauseWarnings = False
        Case Else
            If Left(arg, 1) = "\" Or Left(arg, 1) = "/" Then
                ShowWarning "Unknown argument: " & arg
            Else
                file = arg
            End If
        End Select
    End If
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

pos = InStrRev(file, "\")
pathFile = fso.GetAbsolutePathName(".")
If Right(pathFile, 1) <> "\" Then pathFile = pathFile & "\"
If pos > 0 Then
    pathFile = pathFile & Mid(file, 1, pos)
End If
file = Mid(file, pos+1)
pos = InStrRev(file, ".")
If pos > 0 Then
    texFile = file
Else
    texFile = file & ".tex"
End If

hadError = False
hadWarning = False

Dim errors, warnings, errorCount, warningCount, fatalError, fatalErrorLine, parentSteps

DoLog ""

pdfExec()

If pauseSuccess Or (hadError And pauseErrors) Or (hadWarning And pauseWarnings) Then
    Wscript.StdOut.Write vbCrLf & "[press return]"
    Do While Not WScript.StdIn.AtEndOfLine
       Input = WScript.StdIn.Read(1)
    Loop
End If
If showSuccess And Not hadError and Not hadWarning Then
    ShowMessage "PDF file succesfully created:" & vbcrlf & vbcrlf & "    " & pathFile & pdfFile
End If

If sleep > 0 Then
    DoLog ""
    oShell.AppActivate "PDFExec"
    DoProgressStart "Exiting; press [pause] to interrupt.  "
    sec = sleep
    While sec > 0
        DoProgressStart(sec & ".. ")
        wscript.Sleep(1000)
        sec = sec - 1
    Wend
End If
wscript.quit
