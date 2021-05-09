Imports System.Threading
Imports System.Text
Imports System.IO
Imports Microsoft.Win32
'Imports Microsoft.

Public Class Form1


    Function getSealFromSSHNew(ByVal hostName As String, ByVal userName As String, ByVal password As String, ByVal fecha As Date, ByRef errCodeInt As Integer, ByRef errMessageStr As String) As String


        Dim ssh As New Chilkat.Ssh()
        Dim port As Integer
        Dim channelNum As Integer
        Dim termType As String = "xterm"
        Dim widthInChars As Integer = 80
        Dim heightInChars As Integer = 24
        Dim pixWidth As Integer = 0
        Dim pixHeight As Integer = 0
        Dim cmdOutputStr As String = ""
        Dim msgError As String = ""
        Dim fileNameLog, yearStr, monthStr, dayStr, commandStr As String
        Dim n As Integer
        Dim pollTimeoutMs As Integer = 2000

        Dim success As Boolean = ssh.UnlockComponent("Anything for 30-day trial")

        If (success <> True) Then
            errCodeInt = -1
            errMessageStr = "Error related to component license: " & ssh.LastErrorText
            'MsgBox("Se presento un error #1. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End If

        port = 22

        success = ssh.Connect(hostName, port)
        If (success <> True) Then
            errCodeInt = -2
            errMessageStr = "Error related to the conection: " & ssh.LastErrorText

            'MsgBox("No Funciono Conneccion. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono coneccion")
            Exit Function
            'Else
            '    MsgBox("EXITO Coneccion", MsgBoxStyle.OkOnly, "Funciono Conneccion")
        End If


        success = ssh.AuthenticatePw(userName, password)
        If (success <> True) Then
            errCodeInt = -3
            errMessageStr = "Error related to the authentication: " & ssh.LastErrorText

            'MsgBox("No Funciono Autenticacion. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Autenticacion")
            Exit Function
            'Else
            '    MsgBox("Funciono Autenticacion", MsgBoxStyle.OkOnly, "Funciono Autenticacion. ")
        End If

        channelNum = ssh.OpenSessionChannel()
        If (channelNum < 0) Then
            errCodeInt = -4
            errMessageStr = "Error related to OpenSessionChannel: " & ssh.LastErrorText

            'MsgBox("No Funciono Apertura Canal. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Apertura del canal")
            Exit Function
            'Else
            '    MsgBox("Funciono OpenSessionChannel", MsgBoxStyle.OkOnly, "Funciono OpenSessionChannel. ")

        End If

        success = ssh.SendReqPty(channelNum, termType, widthInChars, heightInChars, pixWidth, pixHeight)
        If (success <> True) Then
            errCodeInt = -5
            errMessageStr = "Error related to SendReqPty: " & ssh.LastErrorText
            'MsgBox("NO funciono Sendreqpty " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Sendreqpty. ")
            Exit Function
            'Else
            '    MsgBox("Funciono SendReqPty", MsgBoxStyle.OkOnly, "Funciono SendReqPty. ")

        End If

        success = ssh.SendReqShell(channelNum)
        If (success <> True) Then
            errCodeInt = -6
            errMessageStr = "Error related to SendReqShell: " & ssh.LastErrorText
            'MsgBox("NO funciono SendReqShell " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono SendReqShell. ")
            Exit Function
            'Else
            '    MsgBox("Funciono SendReqShell", MsgBoxStyle.OkOnly, "Funciono SendReqShell. ")
        End If


        errCodeInt = SSHUntilMatch(ssh, "Enter Selection:", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Function
            'Else
            '    MsgBox("FUNCIONO: ", MsgBoxStyle.OkOnly, "FUNCIONO.")

        End If

        errCodeInt = sentStringToSSHUntilMatch(ssh, "98", "return to the menu):", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Function
            'Else
            '    MsgBox("FUNCIONO: 10 ", MsgBoxStyle.OkOnly, "FUNCIONO.")
        End If

        errCodeInt = sentStringToSSHUntilMatch(ssh, "sudo su - prosys", "Password:", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Function
            'Else
            '    MsgBox("FUNCIONO: 01 ", MsgBoxStyle.OkOnly, "FUNCIONO.")

        End If



        errCodeInt = sentStringToSSH(ssh, password, channelNum, pollTimeoutMs, cmdOutputStr, msgError)

        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Function
        End If


        'errCodeInt = sentStringToSSH(ssh, "df", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        errCodeInt = sentStringToSSH(ssh, "cd elog/files", channelNum, pollTimeoutMs, cmdOutputStr, msgError)

        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Function
        End If


        monthStr = CStr(fecha.Month).PadLeft(2, "0")
        dayStr = CStr(fecha.Day).PadLeft(2, "0")
        yearStr = CStr(fecha.Year).PadLeft(4, "0")

        fileNameLog = "elf" & yearStr & monthStr & dayStr & ".fil"

        commandStr = "grep Seal " & fileNameLog

        errCodeInt = sentStringToSSH(ssh, commandStr, channelNum, pollTimeoutMs, cmdOutputStr, msgError)

        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Function
        End If


        ssh.Disconnect()

        Return cmdOutputStr


    End Function


    Sub getHostModeNew(ByVal hostName As String, ByVal userName As String, ByVal password As String, ByVal hostNumberArray() As Integer, ByRef errCodeInt As Integer, ByRef errMessageStr As String)


        Dim ssh As New Chilkat.Ssh()
        Dim port As Integer
        Dim channelNum, posInt, numInt, isIcludedInt, spareInt As Integer
        Dim termType As String = "xterm"
        Dim widthInChars As Integer = 80
        Dim heightInChars As Integer = 24
        Dim pixWidth As Integer = 0
        Dim pixHeight As Integer = 0
        Dim cmdOutputStr As String = ""
        Dim msgError As String = ""
        Dim n As Integer
        Dim pollTimeoutMs As Integer = 2000

        Dim success As Boolean = ssh.UnlockComponent("Anything for 30-day trial")

        If (success <> True) Then
            errCodeInt = -1
            errMessageStr = "Error related to component license: " & ssh.LastErrorText
            'MsgBox("Se presento un error #1. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "Error")
            Exit Sub
        End If

        port = 22

        success = ssh.Connect(hostName, port)
        If (success <> True) Then
            errCodeInt = -2
            errMessageStr = "Error related to the conection: " & ssh.LastErrorText

            'MsgBox("No Funciono Conneccion. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono coneccion")
            Exit Sub
            'Else
            '    MsgBox("EXITO Coneccion", MsgBoxStyle.OkOnly, "Funciono Conneccion")
        End If


        success = ssh.AuthenticatePw(userName, password)
        If (success <> True) Then
            errCodeInt = -3
            errMessageStr = "Error related to the authentication: " & ssh.LastErrorText

            'MsgBox("No Funciono Autenticacion. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Autenticacion")
            Exit Sub
            'Else
            '    MsgBox("Funciono Autenticacion", MsgBoxStyle.OkOnly, "Funciono Autenticacion. ")
        End If

        channelNum = ssh.OpenSessionChannel()
        If (channelNum < 0) Then
            errCodeInt = -4
            errMessageStr = "Error related to OpenSessionChannel: " & ssh.LastErrorText

            'MsgBox("No Funciono Apertura Canal. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Apertura del canal")
            Exit Sub
            'Else
            '    MsgBox("Funciono OpenSessionChannel", MsgBoxStyle.OkOnly, "Funciono OpenSessionChannel. ")

        End If

        success = ssh.SendReqPty(channelNum, termType, widthInChars, heightInChars, pixWidth, pixHeight)
        If (success <> True) Then
            errCodeInt = -5
            errMessageStr = "Error related to SendReqPty: " & ssh.LastErrorText
            'MsgBox("NO funciono Sendreqpty " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Sendreqpty. ")
            Exit Sub
            'Else
            '    MsgBox("Funciono SendReqPty", MsgBoxStyle.OkOnly, "Funciono SendReqPty. ")

        End If

        success = ssh.SendReqShell(channelNum)
        If (success <> True) Then
            errCodeInt = -6
            errMessageStr = "Error related to SendReqShell: " & ssh.LastErrorText
            'MsgBox("NO funciono SendReqShell " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono SendReqShell. ")
            Exit Sub
            'Else
            '    MsgBox("Funciono SendReqShell", MsgBoxStyle.OkOnly, "Funciono SendReqShell. ")
        End If


        errCodeInt = SSHUntilMatch(ssh, "Enter Selection:", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Sub
            'Else
            '    MsgBox("FUNCIONO: ", MsgBoxStyle.OkOnly, "FUNCIONO.")

        End If

        errCodeInt = sentStringToSSHUntilMatch(ssh, "10", "Enter Selection:", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Sub
            'Else
            '    MsgBox("FUNCIONO: 10 ", MsgBoxStyle.OkOnly, "FUNCIONO.")
        End If

        errCodeInt = sentStringToSSHUntilMatch(ssh, "01", "Selection:", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Sub
            'Else
            '    MsgBox("FUNCIONO: 01 ", MsgBoxStyle.OkOnly, "FUNCIONO.")


        End If

        errCodeInt = sentStringToSSH(ssh, "PEERS", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            'MsgBox(msgError, MsgBoxStyle.OkOnly, "Error")
            Exit Sub
            'Else
            '    RichTextBox1.Text = cmdOutputStr
        End If

        ssh.Disconnect()


        posInt = InStr(1, cmdOutputStr, "live_idn:")
        numInt = CInt(Mid(cmdOutputStr, posInt + 10, 2))
        hostNumberArray(0) = numInt + 1

        posInt = InStr(1, cmdOutputStr, "backup_idn:")
        numInt = CInt(Mid(cmdOutputStr, posInt + 11, 2))
        hostNumberArray(1) = numInt + 1

        posInt = InStr(1, cmdOutputStr, "wait_for_spare:")
        numInt = CInt(Mid(cmdOutputStr, posInt + 16, 2))
        spareInt = numInt + 1

        If ((spareInt = hostNumberArray(0)) Or (spareInt = hostNumberArray(1))) Then

            For i = 1 To 5
                isIcludedInt = Array.IndexOf(hostNumberArray, i)
                If isIcludedInt < 0 Then
                    hostNumberArray(2) = i
                    Exit For
                End If
            Next
        Else
            hostNumberArray(2) = spareInt

        End If


        For i = 1 To 5
            isIcludedInt = Array.IndexOf(hostNumberArray, i)
            If isIcludedInt < 0 Then
                hostNumberArray(3) = i
                Exit For
            End If
        Next

        For i = 1 To 5
            isIcludedInt = Array.IndexOf(hostNumberArray, i)
            If isIcludedInt < 0 Then
                hostNumberArray(4) = i
                Exit For
            End If
        Next

    End Sub



    Sub getHostMode(ByVal hostName As String, ByVal userName As String, ByVal password As String, ByVal hostNumberArray() As Integer, ByRef errCodeInt As Integer, ByRef errMessageStr As String)


        Dim ssh As New Chilkat.Ssh()
        Dim port As Integer
        Dim channelNum, posInt, numInt, isIcludedInt, spareInt As Integer
        Dim termType As String = "xterm"
        Dim widthInChars As Integer = 80
        Dim heightInChars As Integer = 24
        Dim pixWidth As Integer = 0
        Dim pixHeight As Integer = 0
        Dim cmdOutputStr As String = ""
        Dim msgError As String = ""
        Dim n As Integer
        Dim pollTimeoutMs As Integer = 2000

        Dim success As Boolean = ssh.UnlockComponent("Anything for 30-day trial")

        If (success <> True) Then
            errCodeInt = -1
            errMessageStr = "Error related to component license: " & ssh.LastErrorText
            'MsgBox("Se presento un error #1. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "Error")
            Exit Sub
        End If

        port = 22

        success = ssh.Connect(hostName, port)
        If (success <> True) Then
            errCodeInt = -2
            errMessageStr = "Error related to the conection: " & ssh.LastErrorText

            'MsgBox("No Funciono Conneccion. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono coneccion")
            Exit Sub
            'Else
            '    MsgBox("EXITO Coneccion", MsgBoxStyle.OkOnly, "Funciono Conneccion")
        End If


        success = ssh.AuthenticatePw(userName, password)
        If (success <> True) Then
            errCodeInt = -3
            errMessageStr = "Error related to the authentication: " & ssh.LastErrorText

            'MsgBox("No Funciono Autenticacion. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Autenticacion")
            Exit Sub
            'Else
            '    MsgBox("Funciono Autenticacion", MsgBoxStyle.OkOnly, "Funciono Autenticacion. ")
        End If



        channelNum = ssh.OpenSessionChannel()
        If (channelNum < 0) Then
            errCodeInt = -4
            errMessageStr = "Error related to OpenSessionChannel: " & ssh.LastErrorText

            'MsgBox("No Funciono Apertura Canal. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Apertura del canal")
            Exit Sub
        End If



        success = ssh.SendReqPty(channelNum, termType, widthInChars, heightInChars, pixWidth, pixHeight)
        If (success <> True) Then
            errCodeInt = -5
            errMessageStr = "Error related to SendReqPty: " & ssh.LastErrorText
            'MsgBox("NO funciono Sendreqpty " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Sendreqpty. ")
            Exit Sub
        End If



        success = ssh.SendReqShell(channelNum)
        If (success <> True) Then
            errCodeInt = -6
            errMessageStr = "Error related to SendReqShell: " & ssh.LastErrorText
            'MsgBox("NO funciono SendReqShell " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono SendReqShell. ")
            Exit Sub
        End If



        n = ssh.ChannelReadAndPoll(channelNum, pollTimeoutMs)
        If (n < 0) Then
            errCodeInt = -7
            errMessageStr = "Error related to ChannelReadAndPoll: " & ssh.LastErrorText
            'MsgBox("NO funciono ChannelReadAndPoll " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono ChannelReadAndPoll. ")
            Exit Sub
        End If



        cmdOutputStr = ssh.GetReceivedText(channelNum, "ansi")
        If (ssh.LastMethodSuccess <> True) Then
            errCodeInt = -8
            errMessageStr = "Error related to GetReceivedText: " & ssh.LastErrorText
            'MsgBox("NO funciono GetReceivedText " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono GetReceivedText. ")
            Exit Sub
            'Else
            '    RichTextBox1.Text = cmdOutputStr
        End If



        errCodeInt = sentStringToSSH(ssh, "cd gtms/bin", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            'MsgBox(msgError, MsgBoxStyle.OkOnly, "Error")
            Exit Sub
            'Else
            '    RichTextBox1.Text = cmdOutputStr
        End If


        errCodeInt = sentStringToSSH(ssh, "gxvision", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            'MsgBox(msgError, MsgBoxStyle.OkOnly, "Error")
            Exit Sub
            'Else
            '    RichTextBox1.Text = cmdOutputStr
        End If

        errCodeInt = sentStringToSSH(ssh, "PEERS", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            'MsgBox(msgError, MsgBoxStyle.OkOnly, "Error")
            Exit Sub
            'Else
            '    RichTextBox1.Text = cmdOutputStr
        End If

        ssh.Disconnect()


        posInt = InStr(1, cmdOutputStr, "live_idn:")
        numInt = CInt(Mid(cmdOutputStr, posInt + 10, 2))
        hostNumberArray(0) = numInt + 1

        posInt = InStr(1, cmdOutputStr, "backup_idn:")
        numInt = CInt(Mid(cmdOutputStr, posInt + 11, 2))
        hostNumberArray(1) = numInt + 1

        posInt = InStr(1, cmdOutputStr, "wait_for_spare:")
        numInt = CInt(Mid(cmdOutputStr, posInt + 16, 2))
        spareInt = numInt + 1

        If ((spareInt = hostNumberArray(0)) Or (spareInt = hostNumberArray(1))) Then

            For i = 1 To 5
                isIcludedInt = Array.IndexOf(hostNumberArray, i)
                If isIcludedInt < 0 Then
                    hostNumberArray(2) = i
                    Exit For
                End If
            Next
        Else
            hostNumberArray(2) = spareInt

        End If


        For i = 1 To 5
            isIcludedInt = Array.IndexOf(hostNumberArray, i)
            If isIcludedInt < 0 Then
                hostNumberArray(3) = i
                Exit For
            End If
        Next

        For i = 1 To 5
            isIcludedInt = Array.IndexOf(hostNumberArray, i)
            If isIcludedInt < 0 Then
                hostNumberArray(4) = i
                Exit For
            End If
        Next




    End Sub

    Function IsPrimaryFromSSH(ByVal hostName As String, ByVal userName As String, ByVal password As String, ByVal fecha As Date, ByRef errCodeInt As Integer, ByRef errMessageStr As String, ByVal numberHostInt As Integer) As Boolean


        Dim ssh As New Chilkat.Ssh()
        Dim port As Integer
        Dim channelNum As Integer
        Dim termType As String = "xterm"
        Dim widthInChars As Integer = 80
        Dim heightInChars As Integer = 24
        Dim pixWidth As Integer = 0
        Dim pixHeight As Integer = 0
        Dim cmdOutputStr As String = ""
        Dim msgError As String = ""
        Dim fileNameLog, yearStr, monthStr, dayStr, commandStr, hourStr, minStr As String
        Dim n As Integer
        Dim pollTimeoutMs As Integer = 2000
        Dim dateNow As DateTime
        Dim file As System.IO.StreamWriter

        Dim success As Boolean = ssh.UnlockComponent("Anything for 30-day trial")

        If (success <> True) Then
            errCodeInt = -1
            errMessageStr = "Error related to component license: " & ssh.LastErrorText
            'MsgBox("Se presento un error #1. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End If

        port = 22

        success = ssh.Connect(hostName, port)
        If (success <> True) Then
            errCodeInt = -2
            errMessageStr = "Error related to the conection: " & ssh.LastErrorText

            'MsgBox("No Funciono Conneccion. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono coneccion")
            Exit Function
            'Else
            '    MsgBox("EXITO Coneccion", MsgBoxStyle.OkOnly, "Funciono Conneccion")
        End If


        success = ssh.AuthenticatePw(userName, password)
        If (success <> True) Then
            errCodeInt = -3
            errMessageStr = "Error related to the authentication: " & ssh.LastErrorText

            'MsgBox("No Funciono Autenticacion. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Autenticacion")
            Exit Function
            'Else
            '    MsgBox("Funciono Autenticacion", MsgBoxStyle.OkOnly, "Funciono Autenticacion. ")
        End If

        channelNum = ssh.OpenSessionChannel()
        If (channelNum < 0) Then
            errCodeInt = -4
            errMessageStr = "Error related to OpenSessionChannel: " & ssh.LastErrorText

            'MsgBox("No Funciono Apertura Canal. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Apertura del canal")
            Exit Function
        End If

        success = ssh.SendReqPty(channelNum, termType, widthInChars, heightInChars, pixWidth, pixHeight)
        If (success <> True) Then
            errCodeInt = -5
            errMessageStr = "Error related to SendReqPty: " & ssh.LastErrorText
            'MsgBox("NO funciono Sendreqpty " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Sendreqpty. ")
            Exit Function
        End If

        success = ssh.SendReqShell(channelNum)
        If (success <> True) Then
            errCodeInt = -6
            errMessageStr = "Error related to SendReqShell: " & ssh.LastErrorText
            'MsgBox("NO funciono SendReqShell " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono SendReqShell. ")
            Exit Function
        End If


        n = ssh.ChannelReadAndPoll(channelNum, pollTimeoutMs)
        If (n < 0) Then
            errCodeInt = -7
            errMessageStr = "Error related to ChannelReadAndPoll: " & ssh.LastErrorText
            'MsgBox("NO funciono ChannelReadAndPoll " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono ChannelReadAndPoll. ")
            Exit Function
        End If

        cmdOutputStr = ssh.GetReceivedText(channelNum, "utf-8")
        If (ssh.LastMethodSuccess <> True) Then
            errCodeInt = -8
            errMessageStr = "Error related to GetReceivedText: " & ssh.LastErrorText
            'MsgBox("NO funciono GetReceivedText " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono GetReceivedText. ")
            Exit Function
            'Else
            '    RichTextBox1.Text = cmdOutputStr
        End If

        'errCodeInt = sentStringToSSH(ssh, "df", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        errCodeInt = sentStringToSSH(ssh, "cd /oltp/platform/elog/files", channelNum, pollTimeoutMs, cmdOutputStr, msgError)

        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Function
        End If


        monthStr = CStr(fecha.Month).PadLeft(2, "0")
        dayStr = CStr(fecha.Day).PadLeft(2, "0")
        yearStr = CStr(fecha.Year).PadLeft(4, "0")
        'hourStr = CStr(fecha.Hour).PadLeft(2, "0")
        'minStr = CStr(fecha.Minute).PadLeft(2, "0")

        dateNow = Now

        hourStr = CStr(dateNow.Hour).PadLeft(2, "0")
        minStr = Mid(CStr(dateNow.Minute).PadLeft(2, "0"), 1, 1)

        fileNameLog = "elf" & yearStr & monthStr & dayStr & ".fil"

        Select Case numberHostInt
            Case 1
                'grep IL3 elf20180514.fil | grep -c "09:4"

                commandStr = "grep IL1 " & fileNameLog & " | grep -c " & "' " & hourStr & ":" & minStr & "'"
                errCodeInt = sentStringToSSH(ssh, commandStr, channelNum, pollTimeoutMs, cmdOutputStr, msgError)

                If errCodeInt <> 0 Then
                    errCodeInt = -9
                    errMessageStr = ssh.LastErrorText
                    Exit Function
                Else
                    ssh.Disconnect()
                    File = My.Computer.FileSystem.OpenTextFileWriter("IL.txt", False)
                    File.WriteLine(cmdOutputStr)
                    File.Close()
                    'MsgBox(cmdOutputStr)
                    If CInt(cmdOutputStr) > 0 Then
                        Return True
                    Else
                        Return False
                    End If
                End If

            Case 3

                commandStr = "grep IL2 " & fileNameLog & " | grep -c " & hourStr & ":" & minStr
                errCodeInt = sentStringToSSH(ssh, commandStr, channelNum, pollTimeoutMs, cmdOutputStr, msgError)

                If errCodeInt <> 0 Then
                    errCodeInt = -9
                    errMessageStr = ssh.LastErrorText
                    Exit Function
                Else
                    ssh.Disconnect()
                    MsgBox(cmdOutputStr)
                    If CInt(cmdOutputStr) > 0 Then
                        Return True
                    Else
                        Return False
                    End If
                End If

            Case 5

                commandStr = "grep IL3 " & fileNameLog & " | grep -c " & hourStr & ":" & minStr
                errCodeInt = sentStringToSSH(ssh, commandStr, channelNum, pollTimeoutMs, cmdOutputStr, msgError)

                If errCodeInt <> 0 Then
                    errCodeInt = -9
                    errMessageStr = ssh.LastErrorText
                    Exit Function
                Else
                    ssh.Disconnect()
                    If CInt(cmdOutputStr) > 0 Then
                        Return True
                    Else
                        Return False
                    End If
                End If


            Case 7

                commandStr = "grep IL4 " & fileNameLog & " | grep -c " & hourStr & ":" & minStr
                errCodeInt = sentStringToSSH(ssh, commandStr, channelNum, pollTimeoutMs, cmdOutputStr, msgError)

                If errCodeInt <> 0 Then
                    errCodeInt = -9
                    errMessageStr = ssh.LastErrorText
                    Exit Function
                Else
                    ssh.Disconnect()
                    If CInt(cmdOutputStr) > 0 Then
                        Return True
                    Else
                        Return False
                    End If
                End If

            Case 9

                commandStr = "grep IL5 " & fileNameLog & " | grep -c " & hourStr & ":" & minStr
                errCodeInt = sentStringToSSH(ssh, commandStr, channelNum, pollTimeoutMs, cmdOutputStr, msgError)

                If errCodeInt <> 0 Then
                    errCodeInt = -9
                    errMessageStr = ssh.LastErrorText
                    Exit Function
                Else
                    ssh.Disconnect()
                    If CInt(cmdOutputStr) > 0 Then
                        Return True
                    Else
                        Return False
                    End If
                End If

        End Select


    End Function

    
    Function getSealFromSSH(ByVal hostName As String, ByVal userName As String, ByVal password As String, ByVal fecha As Date, ByRef errCodeInt As Integer, ByRef errMessageStr As String) As String


        Dim ssh As New Chilkat.Ssh()
        Dim port As Integer
        Dim channelNum As Integer
        Dim termType As String = "xterm"
        Dim widthInChars As Integer = 80
        Dim heightInChars As Integer = 24
        Dim pixWidth As Integer = 0
        Dim pixHeight As Integer = 0
        Dim cmdOutputStr As String = ""
        Dim msgError As String = ""
        Dim fileNameLog, yearStr, monthStr, dayStr, commandStr As String
        Dim n As Integer
        Dim pollTimeoutMs As Integer = 2000

        Dim success As Boolean = ssh.UnlockComponent("Anything for 30-day trial")

        If (success <> True) Then
            errCodeInt = -1
            errMessageStr = "Error related to component license: " & ssh.LastErrorText
            'MsgBox("Se presento un error #1. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End If

        port = 22

        success = ssh.Connect(hostName, port)
        If (success <> True) Then
            errCodeInt = -2
            errMessageStr = "Error related to the conection: " & ssh.LastErrorText

            'MsgBox("No Funciono Conneccion. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono coneccion")
            Exit Function
            'Else
            '    MsgBox("EXITO Coneccion", MsgBoxStyle.OkOnly, "Funciono Conneccion")
        End If


        success = ssh.AuthenticatePw(userName, password)
        If (success <> True) Then
            errCodeInt = -3
            errMessageStr = "Error related to the authentication: " & ssh.LastErrorText

            'MsgBox("No Funciono Autenticacion. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Autenticacion")
            Exit Function
            'Else
            '    MsgBox("Funciono Autenticacion", MsgBoxStyle.OkOnly, "Funciono Autenticacion. ")
        End If

        channelNum = ssh.OpenSessionChannel()
        If (channelNum < 0) Then
            errCodeInt = -4
            errMessageStr = "Error related to OpenSessionChannel: " & ssh.LastErrorText

            'MsgBox("No Funciono Apertura Canal. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Apertura del canal")
            Exit Function
        End If

        success = ssh.SendReqPty(channelNum, termType, widthInChars, heightInChars, pixWidth, pixHeight)
        If (success <> True) Then
            errCodeInt = -5
            errMessageStr = "Error related to SendReqPty: " & ssh.LastErrorText
            'MsgBox("NO funciono Sendreqpty " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Sendreqpty. ")
            Exit Function
        End If

        success = ssh.SendReqShell(channelNum)
        If (success <> True) Then
            errCodeInt = -6
            errMessageStr = "Error related to SendReqShell: " & ssh.LastErrorText
            'MsgBox("NO funciono SendReqShell " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono SendReqShell. ")
            Exit Function
        End If


        n = ssh.ChannelReadAndPoll(channelNum, pollTimeoutMs)
        If (n < 0) Then
            errCodeInt = -7
            errMessageStr = "Error related to ChannelReadAndPoll: " & ssh.LastErrorText
            'MsgBox("NO funciono ChannelReadAndPoll " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono ChannelReadAndPoll. ")
            Exit Function
        End If

        cmdOutputStr = ssh.GetReceivedText(channelNum, "utf-8")
        If (ssh.LastMethodSuccess <> True) Then
            errCodeInt = -8
            errMessageStr = "Error related to GetReceivedText: " & ssh.LastErrorText
            'MsgBox("NO funciono GetReceivedText " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono GetReceivedText. ")
            Exit Function
            'Else
            '    RichTextBox1.Text = cmdOutputStr
        End If

        'errCodeInt = sentStringToSSH(ssh, "df", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        errCodeInt = sentStringToSSH(ssh, "cd /oltp/platform/elog/files", channelNum, pollTimeoutMs, cmdOutputStr, msgError)

        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Function
        End If


        monthStr = CStr(fecha.Month).PadLeft(2, "0")
        dayStr = CStr(fecha.Day).PadLeft(2, "0")
        yearStr = CStr(fecha.Year).PadLeft(4, "0")

        fileNameLog = "elf" & yearStr & monthStr & dayStr & ".fil"

        commandStr = "grep 'Wag-Seal Original' " & fileNameLog

        errCodeInt = sentStringToSSH(ssh, commandStr, channelNum, pollTimeoutMs, cmdOutputStr, msgError)

        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Function
        End If


        ssh.Disconnect()

        Return cmdOutputStr


    End Function

    Function sentStringToSSH(ByVal ssh As Chilkat.Ssh, ByVal strText As String, ByVal channelNum As Integer, ByVal pollTimeoutMs As Integer, ByRef cmdOutputStr As String, ByRef msgError As String) As Integer
        Dim success As Boolean
        Dim n As Integer

        success = ssh.ChannelSendString(channelNum, strText & vbCrLf, "utf-8")
        If (success <> True) Then
            msgError = "ChannelSendString Error: " + ssh.LastErrorText
            Return -1
        End If

        n = ssh.ChannelReadAndPoll(channelNum, pollTimeoutMs)
        If (n < 0) Then
            msgError = "ChannelReadAndPoll Error: " + ssh.LastErrorText
            Return -2
        End If

        cmdOutputStr = ssh.GetReceivedText(channelNum, "utf-8")
        If (ssh.LastMethodSuccess <> True) Then
            msgError = "GetReceivedText Error: " + ssh.LastErrorText
            Return -3
        Else
            Return 0
        End If

    End Function


    Private Function readFileSeal(ByVal fileName As String, ByVal productStr As String, ByVal drawStr As String, ByRef msgError As String) As String


        Dim reader1 As StreamReader
        Dim strLine, strValue As String
        Dim posInt As Integer

        strValue = ""
        Try
            reader1 = New StreamReader(fileName, Encoding.UTF7)
            strLine = ""

            Do While Not (strLine Is Nothing)
                strLine = reader1.ReadLine()

                If Not (strLine Is Nothing) Then
                    If strLine.Contains(productStr) And strLine.Contains(drawStr) Then
                        posInt = InStrRev(strLine, ":")
                        strValue = Trim((Mid(strLine, posInt + 7)))
                        Exit Do
                    End If
                End If

            Loop
            reader1.Close()
            Return strValue
        Catch ex As Exception
            MsgBox("Error reading file. " & ex.Message, MsgBoxStyle.Information, "Error")
            Return strValue
        End Try
    End Function

    
    
    Private Sub okBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles okBtn.Click

        Dim errCodeInt As Integer = 0
        Dim errMessageStr As String = ""
        Dim file As System.IO.StreamWriter
        Dim infoStr As String = ""
        Dim xlibro As Microsoft.Office.Interop.Excel.Application
        Dim fecha As Date
        Dim dateInt, primaryInt, i As Integer
        Dim monthStr, dayStr, yearStr, fileNameStr, strPathFile, strFilePathReg, dayNameStr, strPathRemoteFile As String
        Dim hostNumberArray(5) As Integer
        Dim conStrinHostStr As String
        Dim substrings() As String
        Dim isPrimaryBool As Boolean


        'substrings(0) -> username
        'substrings(1) -> IP ESTE1
        'substrings(2) -> Password ESTE1
        'substrings(3) -> IP ESTE2
        'substrings(4) -> Password ESTE2
        'substrings(5) -> IP ESTE3
        'substrings(6) -> Password ESTE3
        'substrings(7) -> IP ESTE4
        'substrings(8) -> Password ESTE4
        'substrings(9) -> IP ESTE5
        'substrings(10) -> Password ESTE5

        conStrinHostStr = GetSettingConfigHost(appName, "conStringHost", "").ToString()
        substrings = conStrinHostStr.Split("|")
        fecha = Format(DateTimePicker1.Value, "MM/dd/yyyy")


        Timer1.Enabled = False

        stopBtn.Visible = False
        ProgressBar1.Visible = False

        okBtn.Visible = False
        cancelBtn.Visible = False

        getHostModeNew(substrings(3), substrings(0), substrings(4), hostNumberArray, errCodeInt, errMessageStr)

        If errCodeInt = -3 Then
            i = 1
            isPrimaryBool = False
            While Not (isPrimaryBool) And i <= 9
                isPrimaryBool = IsPrimaryFromSSH(substrings(i), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr, i)
                If isPrimaryBool Then
                    Select Case i
                        Case 1
                            primaryInt = 1
                            hostNumberArray(0) = primaryInt
                            hostNumberArray(1) = 2
                            hostNumberArray(2) = 3
                            hostNumberArray(3) = 4
                            hostNumberArray(4) = 5

                        Case 3
                            primaryInt = 2
                            hostNumberArray(0) = primaryInt
                            hostNumberArray(1) = 1
                            hostNumberArray(2) = 3
                            hostNumberArray(3) = 4
                            hostNumberArray(4) = 5

                        Case 5
                            primaryInt = 3
                            hostNumberArray(0) = primaryInt
                            hostNumberArray(1) = 1
                            hostNumberArray(2) = 2
                            hostNumberArray(3) = 4
                            hostNumberArray(4) = 5

                        Case 7
                            primaryInt = 4
                            hostNumberArray(0) = primaryInt
                            hostNumberArray(1) = 5
                            hostNumberArray(2) = 1
                            hostNumberArray(3) = 2
                            hostNumberArray(4) = 3

                        Case 9
                            primaryInt = 5
                            hostNumberArray(0) = primaryInt
                            hostNumberArray(1) = 4
                            hostNumberArray(2) = 1
                            hostNumberArray(3) = 2
                            hostNumberArray(4) = 3

                    End Select
                Else
                    i += 2
                End If
            End While
            'Mandar coreo informando que hubo un problema de autenticacion
        End If


        'getHostMode(substrings(3), "xfer", "Welcome1", hostNumberArray, errCodeInt, errMessageStr)

        Me.RichTextBox1.Text = "Getting Checksum from Primary..." & vbCrLf
        Me.Refresh()

        Select Case hostNumberArray(0)
            Case 1
                'infoStr = getSealFromSSHNew(substrings(1), substrings(0), substrings(2), fecha, errCodeInt, errMessageStr)
                infoStr = getSealFromSSH(substrings(1), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)

                If errCodeInt <> 0 Then
                    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                End If

            Case 2
                'infoStr = getSealFromSSHNew(substrings(3), substrings(0), substrings(4), fecha, errCodeInt, errMessageStr)

                infoStr = getSealFromSSH(substrings(3), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                If errCodeInt <> 0 Then
                    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                End If
            Case 3
                'infoStr = getSealFromSSHNew(substrings(5), substrings(0), substrings(6), fecha, errCodeInt, errMessageStr)
                infoStr = getSealFromSSH(substrings(5), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                If errCodeInt <> 0 Then
                    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                End If
            Case 4
                'infoStr = getSealFromSSHNew(substrings(7), substrings(0), substrings(8), fecha, errCodeInt, errMessageStr)
                infoStr = getSealFromSSH(substrings(7), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                If errCodeInt <> 0 Then
                    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                End If
            Case 5
                'infoStr = getSealFromSSHNew(substrings(9), substrings(0), substrings(10), fecha, errCodeInt, errMessageStr)
                infoStr = getSealFromSSH(substrings(9), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                If errCodeInt <> 0 Then
                    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                End If
        End Select

        file = My.Computer.FileSystem.OpenTextFileWriter("SEAL_Primary.txt", False)
        file.WriteLine(infoStr)
        file.Close()



        Me.RichTextBox1.Text = "Getting Checksum from Secondary..." & vbCrLf
        Me.Refresh()


        Select Case hostNumberArray(1)
            Case 1
                'infoStr = getSealFromSSHNew(substrings(1), substrings(0), substrings(2), fecha, errCodeInt, errMessageStr)
                infoStr = getSealFromSSH(substrings(1), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)

                If errCodeInt <> 0 Then
                    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                End If

            Case 2
                'infoStr = getSealFromSSHNew(substrings(3), substrings(0), substrings(4), fecha, errCodeInt, errMessageStr)

                infoStr = getSealFromSSH(substrings(3), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                If errCodeInt <> 0 Then
                    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                End If
            Case 3
                'infoStr = getSealFromSSHNew(substrings(5), substrings(0), substrings(6), fecha, errCodeInt, errMessageStr)
                infoStr = getSealFromSSH(substrings(5), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                If errCodeInt <> 0 Then
                    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                End If
            Case 4
                'infoStr = getSealFromSSHNew(substrings(7), substrings(0), substrings(8), fecha, errCodeInt, errMessageStr)
                infoStr = getSealFromSSH(substrings(7), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                If errCodeInt <> 0 Then
                    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                End If
            Case 5
                'infoStr = getSealFromSSHNew(substrings(9), substrings(0), substrings(10), fecha, errCodeInt, errMessageStr)
                infoStr = getSealFromSSH(substrings(9), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                If errCodeInt <> 0 Then
                    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                End If
        End Select

        file = My.Computer.FileSystem.OpenTextFileWriter("SEAL_Secondary.txt", False)
        file.WriteLine(infoStr)
        file.Close()


        Me.RichTextBox1.Text = "Getting Checksum from Spare..." & vbCrLf
        Me.Refresh()

        Select Case hostNumberArray(2)
            Case 1
                'infoStr = getSealFromSSHNew(substrings(1), substrings(0), substrings(2), fecha, errCodeInt, errMessageStr)
                infoStr = getSealFromSSH(substrings(1), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)

                If errCodeInt <> 0 Then
                    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                End If

            Case 2
                'infoStr = getSealFromSSHNew(substrings(3), substrings(0), substrings(4), fecha, errCodeInt, errMessageStr)

                infoStr = getSealFromSSH(substrings(3), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                If errCodeInt <> 0 Then
                    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                End If
            Case 3
                'infoStr = getSealFromSSHNew(substrings(5), substrings(0), substrings(6), fecha, errCodeInt, errMessageStr)
                infoStr = getSealFromSSH(substrings(5), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                If errCodeInt <> 0 Then
                    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                End If
            Case 4
                'infoStr = getSealFromSSHNew(substrings(7), substrings(0), substrings(8), fecha, errCodeInt, errMessageStr)
                infoStr = getSealFromSSH(substrings(7), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                If errCodeInt <> 0 Then
                    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                End If
            Case 5
                'infoStr = getSealFromSSHNew(substrings(9), substrings(0), substrings(10), fecha, errCodeInt, errMessageStr)
                infoStr = getSealFromSSH(substrings(9), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                If errCodeInt <> 0 Then
                    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                End If
        End Select

        file = My.Computer.FileSystem.OpenTextFileWriter("SEAL_Spare.txt", False)
        file.WriteLine(infoStr)
        file.Close()



        Me.RichTextBox1.Text = "Getting Checksum from Spare1..." & vbCrLf
        Me.Refresh()

        Select Case hostNumberArray(3)
            Case 1
                'infoStr = getSealFromSSHNew(substrings(1), substrings(0), substrings(2), fecha, errCodeInt, errMessageStr)
                infoStr = getSealFromSSH(substrings(1), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)

                If errCodeInt <> 0 Then
                    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                End If

            Case 2
                'infoStr = getSealFromSSHNew(substrings(3), substrings(0), substrings(4), fecha, errCodeInt, errMessageStr)

                infoStr = getSealFromSSH(substrings(3), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                If errCodeInt <> 0 Then
                    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                End If
            Case 3
                'infoStr = getSealFromSSHNew(substrings(5), substrings(0), substrings(6), fecha, errCodeInt, errMessageStr)
                infoStr = getSealFromSSH(substrings(5), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                If errCodeInt <> 0 Then
                    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                End If
            Case 4
                'infoStr = getSealFromSSHNew(substrings(7), substrings(0), substrings(8), fecha, errCodeInt, errMessageStr)
                infoStr = getSealFromSSH(substrings(7), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                If errCodeInt <> 0 Then
                    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                End If
            Case 5
                'infoStr = getSealFromSSHNew(substrings(9), substrings(0), substrings(10), fecha, errCodeInt, errMessageStr)
                infoStr = getSealFromSSH(substrings(9), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                If errCodeInt <> 0 Then
                    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                End If
        End Select

        file = My.Computer.FileSystem.OpenTextFileWriter("SEAL_Spare1.txt", False)
        file.WriteLine(infoStr)
        file.Close()


        Me.RichTextBox1.Text = "Getting Checksum from Spare2..." & vbCrLf
        Me.Refresh()

        Select Case hostNumberArray(4)
            Case 1
                'infoStr = getSealFromSSHNew(substrings(1), substrings(0), substrings(2), fecha, errCodeInt, errMessageStr)
                infoStr = getSealFromSSH(substrings(1), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)

                If errCodeInt <> 0 Then
                    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                End If

            Case 2
                'infoStr = getSealFromSSHNew(substrings(3), substrings(0), substrings(4), fecha, errCodeInt, errMessageStr)

                infoStr = getSealFromSSH(substrings(3), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                If errCodeInt <> 0 Then
                    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                End If
            Case 3
                'infoStr = getSealFromSSHNew(substrings(5), substrings(0), substrings(6), fecha, errCodeInt, errMessageStr)
                infoStr = getSealFromSSH(substrings(5), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                If errCodeInt <> 0 Then
                    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                End If
            Case 4
                'infoStr = getSealFromSSHNew(substrings(7), substrings(0), substrings(8), fecha, errCodeInt, errMessageStr)
                infoStr = getSealFromSSH(substrings(7), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                If errCodeInt <> 0 Then
                    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                End If
            Case 5
                'infoStr = getSealFromSSHNew(substrings(9), substrings(0), substrings(10), fecha, errCodeInt, errMessageStr)
                infoStr = getSealFromSSH(substrings(9), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                If errCodeInt <> 0 Then
                    MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                End If
        End Select


        file = My.Computer.FileSystem.OpenTextFileWriter("SEAL_Spare2.txt", False)
        file.WriteLine(infoStr)
        file.Close()



        monthStr = CStr(fecha.Month).PadLeft(2, "0")
        dayStr = CStr(fecha.Day).PadLeft(2, "0")
        yearStr = CStr(fecha.Year).PadLeft(4, "0")
        dateInt = Weekday(fecha)
        dayNameStr = WeekdayName(dateInt)

        strFilePathReg = GetSetting("DrawCheck", "UbiFilesDrawCheck", "").ToString()
        fileNameStr = dayNameStr & "_" & monthStr & dayStr & yearStr & "_" & "Check.xlsx"
        strPathFile = System.Windows.Forms.Application.StartupPath
        Me.RichTextBox1.Text = "Filling Excel file..." & vbCrLf
        Me.Refresh()
        xlibro = CreateObject("Excel.Application")

        If Not (My.Computer.FileSystem.FileExists(strFilePathReg & "\" & fileNameStr)) Then

            Select Case dateInt

                Case 1

                    fileNameStr = strPathFile & "\Sunday_SealCheck.xlsx"
                    xlibro.Workbooks.Open(fileNameStr)
                    xlibro.Visible = True
                    xlibro.Sheets("Primary").Select()
                    xlibro.Cells(1, 6) = fecha


                    fillExcelSundaySeal("Primary", xlibro)
                    fillExcelSundaySeal("Secondary", xlibro)
                    fillExcelSundaySeal("Spare", xlibro)
                    fillExcelSundaySeal("Spare1", xlibro)
                    fillExcelSundaySeal("Spare2", xlibro)


                Case 2
                    fileNameStr = strPathFile & "\ThursdayMonday_SealCheck.xlsx"
                    xlibro.Workbooks.Open(fileNameStr)
                    xlibro.Visible = True
                    xlibro.Sheets("Primary").Select()
                    xlibro.Cells(1, 6) = fecha


                    fillExcelThursdayMondaySeal("Primary", xlibro)
                    fillExcelThursdayMondaySeal("Secondary", xlibro)
                    fillExcelThursdayMondaySeal("Spare", xlibro)
                    fillExcelThursdayMondaySeal("Spare1", xlibro)
                    fillExcelThursdayMondaySeal("Spare2", xlibro)

                Case 3

                    fileNameStr = strPathFile & "\TuesdayFriday_SealCheck.xlsx"
                    xlibro.Workbooks.Open(fileNameStr)
                    xlibro.Visible = True
                    xlibro.Sheets("Primary").Select()
                    xlibro.Cells(1, 6) = fecha


                    fillExcelTuesdayFridaySeal("Primary", xlibro)
                    fillExcelTuesdayFridaySeal("Secondary", xlibro)
                    fillExcelTuesdayFridaySeal("Spare", xlibro)
                    fillExcelTuesdayFridaySeal("Spare1", xlibro)
                    fillExcelTuesdayFridaySeal("Spare2", xlibro)

                Case 4

                    fileNameStr = strPathFile & "\SaturdayWednesday_SealCheck.xlsx"
                    xlibro.Workbooks.Open(fileNameStr)
                    xlibro.Visible = True
                    xlibro.Sheets("Primary").Select()
                    xlibro.Cells(1, 6) = fecha


                    fillExcelSaturdayWednesdaySeal("Primary", xlibro)
                    fillExcelSaturdayWednesdaySeal("Secondary", xlibro)
                    fillExcelSaturdayWednesdaySeal("Spare", xlibro)
                    fillExcelSaturdayWednesdaySeal("Spare1", xlibro)
                    fillExcelSaturdayWednesdaySeal("Spare2", xlibro)


                Case 5
                    fileNameStr = strPathFile & "\ThursdayMonday_SealCheck.xlsx"
                    xlibro.Workbooks.Open(fileNameStr)
                    xlibro.Visible = True
                    xlibro.Sheets("Primary").Select()
                    xlibro.Cells(1, 6) = fecha


                    fillExcelThursdayMondaySeal("Primary", xlibro)
                    fillExcelThursdayMondaySeal("Secondary", xlibro)
                    fillExcelThursdayMondaySeal("Spare", xlibro)
                    fillExcelThursdayMondaySeal("Spare1", xlibro)
                    fillExcelThursdayMondaySeal("Spare2", xlibro)


                Case 6

                    fileNameStr = strPathFile & "\TuesdayFriday_SealCheck.xlsx"
                    xlibro.Workbooks.Open(fileNameStr)
                    xlibro.Visible = True
                    xlibro.Sheets("Primary").Select()
                    xlibro.Cells(1, 6) = fecha


                    fillExcelTuesdayFridaySeal("Primary", xlibro)
                    fillExcelTuesdayFridaySeal("Secondary", xlibro)
                    fillExcelTuesdayFridaySeal("Spare", xlibro)
                    fillExcelTuesdayFridaySeal("Spare1", xlibro)
                    fillExcelTuesdayFridaySeal("Spare2", xlibro)


                Case 7

                    fileNameStr = strPathFile & "\SaturdayWednesday_SealCheck.xlsx"
                    xlibro.Workbooks.Open(fileNameStr)
                    xlibro.Visible = True
                    xlibro.Sheets("Primary").Select()
                    xlibro.Cells(1, 6) = fecha


                    fillExcelSaturdayWednesdaySeal("Primary", xlibro)
                    fillExcelSaturdayWednesdaySeal("Secondary", xlibro)
                    fillExcelSaturdayWednesdaySeal("Spare", xlibro)
                    fillExcelSaturdayWednesdaySeal("Spare1", xlibro)
                    fillExcelSaturdayWednesdaySeal("Spare2", xlibro)

            End Select


            xlibro.Sheets("Primary").Select()
            xlibro.Cells(3, 2) = "NYESTE" & hostNumberArray(0)

            xlibro.Sheets("Secondary").Select()
            xlibro.Cells(3, 2) = "NYESTE" & hostNumberArray(1)

            xlibro.Sheets("Spare").Select()
            xlibro.Cells(3, 2) = "NYESTE" & hostNumberArray(2)

            xlibro.Sheets("Spare1").Select()
            xlibro.Cells(3, 2) = "NYESTE" & hostNumberArray(3)

            xlibro.Sheets("Spare2").Select()
            xlibro.Cells(3, 2) = "NYESTE" & hostNumberArray(4)

            xlibro.Sheets("Primary").Select()

            strFilePathReg = GetSetting("DrawCheck", "UbiFilesDrawCheck", "").ToString()

            fileNameStr = dayNameStr & "_" & monthStr & dayStr & yearStr & "_" & "Check.xlsx"

            xlibro.ActiveWorkbook.SaveAs(strFilePathReg & "\" & fileNameStr)

            'xlibro.ActiveWorkbook.Save(strFilePathReg & "\" & fileNameStr)

            xlibro.Workbooks.Close()

            xlibro.Quit()


        Else

            fileNameStr = (strFilePathReg & "\" & fileNameStr)

            Select Case dateInt

                Case 1

                    'fileNameStr = strPathFile & "\Sunday_SealCheck.xlsx"
                    xlibro.Workbooks.Open(fileNameStr)
                    xlibro.Visible = True
                    xlibro.Sheets("Primary").Select()
                    xlibro.Cells(1, 6) = fecha

                    fillExcelSundaySeal("Primary", xlibro)
                    fillExcelSundaySeal("Secondary", xlibro)
                    fillExcelSundaySeal("Spare", xlibro)
                    fillExcelSundaySeal("Spare1", xlibro)
                    fillExcelSundaySeal("Spare2", xlibro)


                Case 2
                    'fileNameStr = strPathFile & "\ThursdayMonday_SealCheck.xlsx"
                    xlibro.Workbooks.Open(fileNameStr)
                    xlibro.Visible = True
                    xlibro.Sheets("Primary").Select()
                    xlibro.Cells(1, 6) = fecha


                    fillExcelThursdayMondaySeal("Primary", xlibro)
                    fillExcelThursdayMondaySeal("Secondary", xlibro)
                    fillExcelThursdayMondaySeal("Spare", xlibro)
                    fillExcelThursdayMondaySeal("Spare1", xlibro)
                    fillExcelThursdayMondaySeal("Spare2", xlibro)

                Case 3

                    'fileNameStr = strPathFile & "\TuesdayFriday_SealCheck.xlsx"
                    xlibro.Workbooks.Open(fileNameStr)
                    xlibro.Visible = True
                    xlibro.Sheets("Primary").Select()
                    xlibro.Cells(1, 6) = fecha


                    fillExcelTuesdayFridaySeal("Primary", xlibro)
                    fillExcelTuesdayFridaySeal("Secondary", xlibro)
                    fillExcelTuesdayFridaySeal("Spare", xlibro)
                    fillExcelTuesdayFridaySeal("Spare1", xlibro)
                    fillExcelTuesdayFridaySeal("Spare2", xlibro)

                Case 4

                    'fileNameStr = strPathFile & "\SaturdayWednesday_SealCheck.xlsx"
                    xlibro.Workbooks.Open(fileNameStr)
                    xlibro.Visible = True
                    xlibro.Sheets("Primary").Select()
                    xlibro.Cells(1, 6) = fecha

                    fillExcelSaturdayWednesdaySeal("Primary", xlibro)
                    fillExcelSaturdayWednesdaySeal("Secondary", xlibro)
                    fillExcelSaturdayWednesdaySeal("Spare", xlibro)
                    fillExcelSaturdayWednesdaySeal("Spare1", xlibro)
                    fillExcelSaturdayWednesdaySeal("Spare2", xlibro)


                Case 5
                    'fileNameStr = strPathFile & "\ThursdayMonday_SealCheck.xlsx"
                    xlibro.Workbooks.Open(fileNameStr)
                    xlibro.Visible = True
                    xlibro.Sheets("Primary").Select()
                    xlibro.Cells(1, 6) = fecha

                    fillExcelThursdayMondaySeal("Primary", xlibro)
                    fillExcelThursdayMondaySeal("Secondary", xlibro)
                    fillExcelThursdayMondaySeal("Spare", xlibro)
                    fillExcelThursdayMondaySeal("Spare1", xlibro)
                    fillExcelThursdayMondaySeal("Spare2", xlibro)


                Case 6

                    'fileNameStr = strPathFile & "\TuesdayFriday_SealCheck.xlsx"
                    xlibro.Workbooks.Open(fileNameStr)
                    xlibro.Visible = True
                    xlibro.Sheets("Primary").Select()
                    xlibro.Cells(1, 6) = fecha

                    fillExcelTuesdayFridaySeal("Primary", xlibro)
                    fillExcelTuesdayFridaySeal("Secondary", xlibro)
                    fillExcelTuesdayFridaySeal("Spare", xlibro)
                    fillExcelTuesdayFridaySeal("Spare1", xlibro)
                    fillExcelTuesdayFridaySeal("Spare2", xlibro)


                Case 7

                    'fileNameStr = strPathFile & "\SaturdayWednesday_SealCheck.xlsx"
                    xlibro.Workbooks.Open(fileNameStr)
                    xlibro.Visible = True
                    xlibro.Sheets("Primary").Select()
                    xlibro.Cells(1, 6) = fecha


                    fillExcelSaturdayWednesdaySeal("Primary", xlibro)
                    fillExcelSaturdayWednesdaySeal("Secondary", xlibro)
                    fillExcelSaturdayWednesdaySeal("Spare", xlibro)
                    fillExcelSaturdayWednesdaySeal("Spare1", xlibro)
                    fillExcelSaturdayWednesdaySeal("Spare2", xlibro)

            End Select


            xlibro.Sheets("Primary").Select()
            xlibro.Cells(3, 2) = "NYESTE" & hostNumberArray(0)

            xlibro.Sheets("Secondary").Select()
            xlibro.Cells(3, 2) = "NYESTE" & hostNumberArray(1)

            xlibro.Sheets("Spare").Select()
            xlibro.Cells(3, 2) = "NYESTE" & hostNumberArray(2)

            xlibro.Sheets("Spare1").Select()
            xlibro.Cells(3, 2) = "NYESTE" & hostNumberArray(3)

            xlibro.Sheets("Spare2").Select()
            xlibro.Cells(3, 2) = "NYESTE" & hostNumberArray(4)

            xlibro.Sheets("Primary").Select()

            'strFilePathReg = GetSetting("DrawCheck", "UbiFilesDrawCheck", "").ToString()

            'fileNameStr = dayNameStr & "_" & monthStr & dayStr & yearStr & "_" & "Check.xlsx"

            'xlibro.ActiveWorkbook.SaveAs(strFilePathReg & "\" & fileNameStr)

            xlibro.ActiveWorkbook.Save()

            xlibro.Workbooks.Close()

            xlibro.Quit()


        End If

        Me.Cursor = Cursors.Default
        Me.RichTextBox1.Text = "All Done."
        Me.Refresh()
        Me.cancelBtn.Enabled = True

    End Sub

    Public Function GetSetting(ByVal APP_NAME As String, ByVal Keyname As String, Optional ByVal DefVal As String = "") As String
        Dim Key As RegistryKey
        Try
            Key = Registry.CurrentUser.CreateSubKey("SOFTWARE").CreateSubKey("IGT").CreateSubKey(APP_NAME)
            Return Key.GetValue(Keyname, DefVal)


        Catch
            Return DefVal
        End Try
    End Function


    Private Sub cancelBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cancelBtn.Click
        Me.Close()
    End Sub

    
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        DateTimePicker1.Value = Now
        ProgressBar1.Maximum = 100
        okBtn.Visible = False
        cancelBtn.Visible = False
        Timer1.Enabled = True
    End Sub

    
    Private Sub fillExcelSundaySeal(ByVal hostName As String, ByRef bookExcel As Microsoft.Office.Interop.Excel.Application)

        Dim drawFindStr, checkSumStr As String
        Dim errMessageStr As String = ""


        bookExcel.Sheets(hostName).Select()
        drawFindStr = bookExcel.Cells(4, 3).Value
        checkSumStr = readFileSeal("SEAL_" & hostName & ".txt", "pck3:drawseal", drawFindStr, errMessageStr)
        bookExcel.Cells(6, 3) = Trim(checkSumStr)

        checkSumStr = readFileSeal("SEAL_" & hostName & ".txt", "pck3:winseal", drawFindStr, errMessageStr)
        bookExcel.Cells(7, 3) = Trim(checkSumStr)


        drawFindStr = bookExcel.Cells(10, 3).Value
        checkSumStr = readFileSeal("SEAL_" & hostName & ".txt", "pck4:drawseal", drawFindStr, errMessageStr)
        bookExcel.Cells(12, 3) = Trim(checkSumStr)

        checkSumStr = readFileSeal("SEAL_" & hostName & ".txt", "pck4:winseal", drawFindStr, errMessageStr)
        bookExcel.Cells(13, 3) = Trim(checkSumStr)


        drawFindStr = bookExcel.Cells(18, 3).Value
        checkSumStr = readFileSeal("SEAL_" & hostName & ".txt", "pck3:drawseal", drawFindStr, errMessageStr)
        bookExcel.Cells(20, 3) = Trim(checkSumStr)

        checkSumStr = readFileSeal("SEAL_" & hostName & ".txt", "pck3:winseal", drawFindStr, errMessageStr)
        bookExcel.Cells(21, 3) = Trim(checkSumStr)


        drawFindStr = bookExcel.Cells(24, 3).Value
        checkSumStr = readFileSeal("SEAL_" & hostName & ".txt", "pck4:drawseal", drawFindStr, errMessageStr)
        bookExcel.Cells(26, 3) = Trim(checkSumStr)

        checkSumStr = readFileSeal("SEAL_" & hostName & ".txt", "pck4:winseal", drawFindStr, errMessageStr)
        bookExcel.Cells(27, 3) = Trim(checkSumStr)


        drawFindStr = bookExcel.Cells(30, 3).Value
        checkSumStr = readFileSeal("SEAL_" & hostName & ".txt", "dkno:drawseal", drawFindStr, errMessageStr)
        bookExcel.Cells(33, 3) = Trim(checkSumStr)

        checkSumStr = readFileSeal("SEAL_" & hostName & ".txt", "dkno:winseal", drawFindStr, errMessageStr)
        bookExcel.Cells(34, 3) = Trim(checkSumStr)


        drawFindStr = bookExcel.Cells(49, 3).Value
        checkSumStr = readFileSeal("SEAL_" & hostName & ".txt", "csh5:drawseal", drawFindStr, errMessageStr)
        bookExcel.Cells(51, 3) = Trim(checkSumStr)

        checkSumStr = readFileSeal("SEAL_" & hostName & ".txt", "csh5:winseal", drawFindStr, errMessageStr)
        bookExcel.Cells(52, 3) = Trim(checkSumStr)


        drawFindStr = bookExcel.Cells(55, 3).Value
        checkSumStr = readFileSeal("SEAL_" & hostName & ".txt", "life:drawseal", drawFindStr, errMessageStr)
        bookExcel.Cells(57, 3) = Trim(checkSumStr)

        checkSumStr = readFileSeal("SEAL_" & hostName & ".txt", "life:winseal", drawFindStr, errMessageStr)
        bookExcel.Cells(58, 3) = Trim(checkSumStr)


        'drawFindStr = bookExcel.Cells(49, 3).Value
        'checkSumStr = readFileSeal("SEAL_" & hostName & ".txt", "csh5:drawseal", drawFindStr, errMessageStr)
        'bookExcel.Cells(51, 3) = Trim(checkSumStr)

        'checkSumStr = readFileSeal("SEAL_" & hostName & ".txt", "csh5:winseal", drawFindStr, errMessageStr)
        'bookExcel.Cells(52, 3) = Trim(checkSumStr)



    End Sub



    Private Sub fillExcelSaturdayWednesdaySeal(ByVal hostName As String, ByRef bookExcel As Microsoft.Office.Interop.Excel.Application)

        Dim drawFindStr, checkSumStr As String
        Dim errMessageStr As String = ""


        fillExcelSundaySeal(hostName, bookExcel)

        drawFindStr = bookExcel.Cells(37, 3).Value
        checkSumStr = readFileSeal("SEAL_" & hostName & ".txt", "pwrb:drawseal", drawFindStr, errMessageStr)
        bookExcel.Cells(39, 3) = Trim(checkSumStr)

        checkSumStr = readFileSeal("SEAL_" & hostName & ".txt", "pwrb:winseal", drawFindStr, errMessageStr)
        bookExcel.Cells(40, 3) = Trim(checkSumStr)


        drawFindStr = bookExcel.Cells(43, 3).Value
        checkSumStr = readFileSeal("SEAL_" & hostName & ".txt", "loto:drawseal", drawFindStr, errMessageStr)
        bookExcel.Cells(45, 3) = Trim(checkSumStr)

        checkSumStr = readFileSeal("SEAL_" & hostName & ".txt", "loto:winseal", drawFindStr, errMessageStr)
        bookExcel.Cells(46, 3) = Trim(checkSumStr)


    End Sub



    Private Sub fillExcelTuesdayFridaySeal(ByVal hostName As String, ByRef bookExcel As Microsoft.Office.Interop.Excel.Application)

        Dim drawFindStr, checkSumStr As String
        Dim errMessageStr As String = ""


        fillExcelSundaySeal(hostName, bookExcel)

        drawFindStr = bookExcel.Cells(61, 3).Value
        checkSumStr = readFileSeal("SEAL_" & hostName & ".txt", "bigg:drawseal", drawFindStr, errMessageStr)
        bookExcel.Cells(63, 3) = Trim(checkSumStr)

        checkSumStr = readFileSeal("SEAL_" & hostName & ".txt", "bigg:winseal", drawFindStr, errMessageStr)
        bookExcel.Cells(64, 3) = Trim(checkSumStr)


    End Sub

    Private Sub fillExcelThursdayMondaySeal(ByVal hostName As String, ByRef bookExcel As Microsoft.Office.Interop.Excel.Application)

        'Dim drawFindStr, checkSumStr As String
        'Dim errMessageStr As String = ""


        fillExcelSundaySeal(hostName, bookExcel)

        'drawFindStr = bookExcel.Cells(55, 3).Value
        'checkSumStr = readFileSeal("SEAL_" & hostName & ".txt", "life:drawseal", drawFindStr, errMessageStr)
        'bookExcel.Cells(57, 3) = Trim(checkSumStr)

        'checkSumStr = readFileSeal("SEAL_" & hostName & ".txt", "life:winseal", drawFindStr, errMessageStr)
        'bookExcel.Cells(58, 3) = Trim(checkSumStr)


    End Sub

    
    Private Sub stopBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles stopBtn.Click
        Timer1.Enabled = False
        stopBtn.Visible = False
        ProgressBar1.Visible = False

        okBtn.Visible = True
        cancelBtn.Visible = True
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim errCodeInt As Integer = 0
        Dim errMessageStr As String = ""
        Dim file As System.IO.StreamWriter
        Dim infoStr As String = ""
        Dim xlibro As Microsoft.Office.Interop.Excel.Application
        Dim fecha, dateNow As DateTime
        Dim dateInt, primaryInt, i As Integer
        Dim monthStr, dayStr, yearStr, fileNameStr, strPathFile, strFilePathReg, dayNameStr, strPathRemoteFile As String
        Dim hostNumberArray(5) As Integer
        Dim conStrinHostStr, hourStr, minStr As String
        Dim substrings() As String
        Dim isPrimaryBool As Boolean

        

        'substrings(0) -> username
        'substrings(1) -> IP ESTE1
        'substrings(2) -> Password ESTE1
        'substrings(3) -> IP ESTE2
        'substrings(4) -> Password ESTE2
        'substrings(5) -> IP ESTE3
        'substrings(6) -> Password ESTE3
        'substrings(7) -> IP ESTE4
        'substrings(8) -> Password ESTE4
        'substrings(9) -> IP ESTE5
        'substrings(10) -> Password ESTE5

        conStrinHostStr = GetSettingConfigHost(appName, "conStringHost", "").ToString()
        substrings = conStrinHostStr.Split("|")
        fecha = Format(DateTimePicker1.Value, "MM/dd/yyyy")

        'dateNow = Now

        'hourStr = CStr(dateNow.Hour).PadLeft(2, "0")
        'minStr = Mid(CStr(dateNow.Minute).PadLeft(2, "0"), 1, 1)

        Static count As Integer = 0
        count = count + 10
        If count <= 100 Then
            ProgressBar1.Value = count
        Else

            Timer1.Enabled = False

            stopBtn.Visible = False
            ProgressBar1.Visible = False

            okBtn.Visible = False
            cancelBtn.Visible = False

            getHostModeNew(substrings(3), substrings(0), substrings(4), hostNumberArray, errCodeInt, errMessageStr)
            If errCodeInt = -3 Then
                i = 1
                isPrimaryBool = False
                While Not (isPrimaryBool) And i <= 9
                    isPrimaryBool = IsPrimaryFromSSH(substrings(i), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr, i)
                    If isPrimaryBool Then
                        Select Case i
                            Case 1
                                primaryInt = 1
                                hostNumberArray(0) = primaryInt
                                hostNumberArray(1) = 2
                                hostNumberArray(2) = 3
                                hostNumberArray(3) = 4
                                hostNumberArray(4) = 5

                            Case 3
                                primaryInt = 2
                                hostNumberArray(0) = primaryInt
                                hostNumberArray(1) = 1
                                hostNumberArray(2) = 3
                                hostNumberArray(3) = 4
                                hostNumberArray(4) = 5

                            Case 5
                                primaryInt = 3
                                hostNumberArray(0) = primaryInt
                                hostNumberArray(1) = 1
                                hostNumberArray(2) = 2
                                hostNumberArray(3) = 4
                                hostNumberArray(4) = 5

                            Case 7
                                primaryInt = 4
                                hostNumberArray(0) = primaryInt
                                hostNumberArray(1) = 5
                                hostNumberArray(2) = 1
                                hostNumberArray(3) = 2
                                hostNumberArray(4) = 3

                            Case 9
                                primaryInt = 5
                                hostNumberArray(0) = primaryInt
                                hostNumberArray(1) = 4
                                hostNumberArray(2) = 1
                                hostNumberArray(3) = 2
                                hostNumberArray(4) = 3

                        End Select
                    Else
                        i += 2
                    End If
                End While
                'Mandar coreo informando que hubo un problema de autenticacion
            End If

            'getHostMode(substrings(3), "xfer", "Welcome1", hostNumberArray, errCodeInt, errMessageStr)


            Me.RichTextBox1.Text = "Getting Checksum from Primary..." & vbCrLf
            Me.Refresh()

            Select Case hostNumberArray(0)
                Case 1
                    'infoStr = getSealFromSSHNew(substrings(1), substrings(0), substrings(2), fecha, errCodeInt, errMessageStr)
                    infoStr = getSealFromSSH(substrings(1), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)

                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If

                Case 2
                    'infoStr = getSealFromSSHNew(substrings(3), substrings(0), substrings(4), fecha, errCodeInt, errMessageStr)

                    infoStr = getSealFromSSH(substrings(3), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If
                Case 3
                    'infoStr = getSealFromSSHNew(substrings(5), substrings(0), substrings(6), fecha, errCodeInt, errMessageStr)
                    infoStr = getSealFromSSH(substrings(5), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If
                Case 4
                    'infoStr = getSealFromSSHNew(substrings(7), substrings(0), substrings(8), fecha, errCodeInt, errMessageStr)
                    infoStr = getSealFromSSH(substrings(7), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If
                Case 5
                    'infoStr = getSealFromSSHNew(substrings(9), substrings(0), substrings(10), fecha, errCodeInt, errMessageStr)
                    infoStr = getSealFromSSH(substrings(9), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If
            End Select

            file = My.Computer.FileSystem.OpenTextFileWriter("SEAL_Primary.txt", False)
            file.WriteLine(infoStr)
            file.Close()



            Me.RichTextBox1.Text = "Getting Checksum from Secondary..." & vbCrLf
            Me.Refresh()


            Select Case hostNumberArray(1)
                Case 1
                    'infoStr = getSealFromSSHNew(substrings(1), substrings(0), substrings(2), fecha, errCodeInt, errMessageStr)
                    infoStr = getSealFromSSH(substrings(1), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)

                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If

                Case 2
                    'infoStr = getSealFromSSHNew(substrings(3), substrings(0), substrings(4), fecha, errCodeInt, errMessageStr)

                    infoStr = getSealFromSSH(substrings(3), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If
                Case 3
                    'infoStr = getSealFromSSHNew(substrings(5), substrings(0), substrings(6), fecha, errCodeInt, errMessageStr)
                    infoStr = getSealFromSSH(substrings(5), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If
                Case 4
                    'infoStr = getSealFromSSHNew(substrings(7), substrings(0), substrings(8), fecha, errCodeInt, errMessageStr)
                    infoStr = getSealFromSSH(substrings(7), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If
                Case 5
                    'infoStr = getSealFromSSHNew(substrings(9), substrings(0), substrings(10), fecha, errCodeInt, errMessageStr)
                    infoStr = getSealFromSSH(substrings(9), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If
            End Select

            file = My.Computer.FileSystem.OpenTextFileWriter("SEAL_Secondary.txt", False)
            file.WriteLine(infoStr)
            file.Close()


            Me.RichTextBox1.Text = "Getting Checksum from Spare..." & vbCrLf
            Me.Refresh()

            Select Case hostNumberArray(2)
                Case 1
                    'infoStr = getSealFromSSHNew(substrings(1), substrings(0), substrings(2), fecha, errCodeInt, errMessageStr)
                    infoStr = getSealFromSSH(substrings(1), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)

                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If

                Case 2
                    'infoStr = getSealFromSSHNew(substrings(3), substrings(0), substrings(4), fecha, errCodeInt, errMessageStr)

                    infoStr = getSealFromSSH(substrings(3), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If
                Case 3
                    'infoStr = getSealFromSSHNew(substrings(5), substrings(0), substrings(6), fecha, errCodeInt, errMessageStr)
                    infoStr = getSealFromSSH(substrings(5), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If
                Case 4
                    'infoStr = getSealFromSSHNew(substrings(7), substrings(0), substrings(8), fecha, errCodeInt, errMessageStr)
                    infoStr = getSealFromSSH(substrings(7), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If
                Case 5
                    'infoStr = getSealFromSSHNew(substrings(9), substrings(0), substrings(10), fecha, errCodeInt, errMessageStr)
                    infoStr = getSealFromSSH(substrings(9), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If
            End Select

            file = My.Computer.FileSystem.OpenTextFileWriter("SEAL_Spare.txt", False)
            file.WriteLine(infoStr)
            file.Close()



            Me.RichTextBox1.Text = "Getting Checksum from Spare1..." & vbCrLf
            Me.Refresh()

            Select Case hostNumberArray(3)
                Case 1
                    'infoStr = getSealFromSSHNew(substrings(1), substrings(0), substrings(2), fecha, errCodeInt, errMessageStr)
                    infoStr = getSealFromSSH(substrings(1), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)

                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If

                Case 2
                    'infoStr = getSealFromSSHNew(substrings(3), substrings(0), substrings(4), fecha, errCodeInt, errMessageStr)

                    infoStr = getSealFromSSH(substrings(3), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If
                Case 3
                    'infoStr = getSealFromSSHNew(substrings(5), substrings(0), substrings(6), fecha, errCodeInt, errMessageStr)
                    infoStr = getSealFromSSH(substrings(5), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If
                Case 4
                    'infoStr = getSealFromSSHNew(substrings(7), substrings(0), substrings(8), fecha, errCodeInt, errMessageStr)
                    infoStr = getSealFromSSH(substrings(7), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If
                Case 5
                    'infoStr = getSealFromSSHNew(substrings(9), substrings(0), substrings(10), fecha, errCodeInt, errMessageStr)
                    infoStr = getSealFromSSH(substrings(9), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If
            End Select

            file = My.Computer.FileSystem.OpenTextFileWriter("SEAL_Spare1.txt", False)
            file.WriteLine(infoStr)
            file.Close()


            Me.RichTextBox1.Text = "Getting Checksum from Spare2..." & vbCrLf
            Me.Refresh()

            Select Case hostNumberArray(4)
                Case 1
                    'infoStr = getSealFromSSHNew(substrings(1), substrings(0), substrings(2), fecha, errCodeInt, errMessageStr)
                    infoStr = getSealFromSSH(substrings(1), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)

                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If

                Case 2
                    'infoStr = getSealFromSSHNew(substrings(3), substrings(0), substrings(4), fecha, errCodeInt, errMessageStr)

                    infoStr = getSealFromSSH(substrings(3), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If
                Case 3
                    'infoStr = getSealFromSSHNew(substrings(5), substrings(0), substrings(6), fecha, errCodeInt, errMessageStr)
                    infoStr = getSealFromSSH(substrings(5), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If
                Case 4
                    'infoStr = getSealFromSSHNew(substrings(7), substrings(0), substrings(8), fecha, errCodeInt, errMessageStr)
                    infoStr = getSealFromSSH(substrings(7), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If
                Case 5
                    'infoStr = getSealFromSSHNew(substrings(9), substrings(0), substrings(10), fecha, errCodeInt, errMessageStr)
                    infoStr = getSealFromSSH(substrings(9), "xfer", "Welcome1", fecha, errCodeInt, errMessageStr)
                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If
            End Select


            file = My.Computer.FileSystem.OpenTextFileWriter("SEAL_Spare2.txt", False)
            file.WriteLine(infoStr)
            file.Close()



            monthStr = CStr(fecha.Month).PadLeft(2, "0")
            dayStr = CStr(fecha.Day).PadLeft(2, "0")
            yearStr = CStr(fecha.Year).PadLeft(4, "0")
            dateInt = Weekday(fecha)
            dayNameStr = WeekdayName(dateInt)

            strFilePathReg = GetSetting("DrawCheck", "UbiFilesDrawCheck", "").ToString()
            fileNameStr = dayNameStr & "_" & monthStr & dayStr & yearStr & "_" & "Check.xlsx"
            strPathFile = System.Windows.Forms.Application.StartupPath
            Me.RichTextBox1.Text = "Filling Excel file..." & vbCrLf
            Me.Refresh()
            xlibro = CreateObject("Excel.Application")

            If Not (My.Computer.FileSystem.FileExists(strFilePathReg & "\" & fileNameStr)) Then

                Select Case dateInt

                    Case 1

                        fileNameStr = strPathFile & "\Sunday_SealCheck.xlsx"
                        xlibro.Workbooks.Open(fileNameStr)
                        xlibro.Visible = False
                        xlibro.Sheets("Primary").Select()
                        xlibro.Cells(1, 6) = fecha


                        fillExcelSundaySeal("Primary", xlibro)
                        fillExcelSundaySeal("Secondary", xlibro)
                        fillExcelSundaySeal("Spare", xlibro)
                        fillExcelSundaySeal("Spare1", xlibro)
                        fillExcelSundaySeal("Spare2", xlibro)


                    Case 2
                        fileNameStr = strPathFile & "\ThursdayMonday_SealCheck.xlsx"
                        xlibro.Workbooks.Open(fileNameStr)
                        xlibro.Visible = False
                        xlibro.Sheets("Primary").Select()
                        xlibro.Cells(1, 6) = fecha


                        fillExcelThursdayMondaySeal("Primary", xlibro)
                        fillExcelThursdayMondaySeal("Secondary", xlibro)
                        fillExcelThursdayMondaySeal("Spare", xlibro)
                        fillExcelThursdayMondaySeal("Spare1", xlibro)
                        fillExcelThursdayMondaySeal("Spare2", xlibro)

                    Case 3

                        fileNameStr = strPathFile & "\TuesdayFriday_SealCheck.xlsx"
                        xlibro.Workbooks.Open(fileNameStr)
                        xlibro.Visible = False
                        xlibro.Sheets("Primary").Select()
                        xlibro.Cells(1, 6) = fecha


                        fillExcelTuesdayFridaySeal("Primary", xlibro)
                        fillExcelTuesdayFridaySeal("Secondary", xlibro)
                        fillExcelTuesdayFridaySeal("Spare", xlibro)
                        fillExcelTuesdayFridaySeal("Spare1", xlibro)
                        fillExcelTuesdayFridaySeal("Spare2", xlibro)

                    Case 4

                        fileNameStr = strPathFile & "\SaturdayWednesday_SealCheck.xlsx"
                        xlibro.Workbooks.Open(fileNameStr)
                        xlibro.Visible = False
                        xlibro.Sheets("Primary").Select()
                        xlibro.Cells(1, 6) = fecha


                        fillExcelSaturdayWednesdaySeal("Primary", xlibro)
                        fillExcelSaturdayWednesdaySeal("Secondary", xlibro)
                        fillExcelSaturdayWednesdaySeal("Spare", xlibro)
                        fillExcelSaturdayWednesdaySeal("Spare1", xlibro)
                        fillExcelSaturdayWednesdaySeal("Spare2", xlibro)


                    Case 5
                        fileNameStr = strPathFile & "\ThursdayMonday_SealCheck.xlsx"
                        xlibro.Workbooks.Open(fileNameStr)
                        xlibro.Visible = False
                        xlibro.Sheets("Primary").Select()
                        xlibro.Cells(1, 6) = fecha


                        fillExcelThursdayMondaySeal("Primary", xlibro)
                        fillExcelThursdayMondaySeal("Secondary", xlibro)
                        fillExcelThursdayMondaySeal("Spare", xlibro)
                        fillExcelThursdayMondaySeal("Spare1", xlibro)
                        fillExcelThursdayMondaySeal("Spare2", xlibro)


                    Case 6

                        fileNameStr = strPathFile & "\TuesdayFriday_SealCheck.xlsx"
                        xlibro.Workbooks.Open(fileNameStr)
                        xlibro.Visible = False
                        xlibro.Sheets("Primary").Select()
                        xlibro.Cells(1, 6) = fecha


                        fillExcelTuesdayFridaySeal("Primary", xlibro)
                        fillExcelTuesdayFridaySeal("Secondary", xlibro)
                        fillExcelTuesdayFridaySeal("Spare", xlibro)
                        fillExcelTuesdayFridaySeal("Spare1", xlibro)
                        fillExcelTuesdayFridaySeal("Spare2", xlibro)


                    Case 7

                        fileNameStr = strPathFile & "\SaturdayWednesday_SealCheck.xlsx"
                        xlibro.Workbooks.Open(fileNameStr)
                        xlibro.Visible = False
                        xlibro.Sheets("Primary").Select()
                        xlibro.Cells(1, 6) = fecha


                        fillExcelSaturdayWednesdaySeal("Primary", xlibro)
                        fillExcelSaturdayWednesdaySeal("Secondary", xlibro)
                        fillExcelSaturdayWednesdaySeal("Spare", xlibro)
                        fillExcelSaturdayWednesdaySeal("Spare1", xlibro)
                        fillExcelSaturdayWednesdaySeal("Spare2", xlibro)

                End Select


                xlibro.Sheets("Primary").Select()
                xlibro.Cells(3, 2) = "NYESTE" & hostNumberArray(0)

                xlibro.Sheets("Secondary").Select()
                xlibro.Cells(3, 2) = "NYESTE" & hostNumberArray(1)

                xlibro.Sheets("Spare").Select()
                xlibro.Cells(3, 2) = "NYESTE" & hostNumberArray(2)

                xlibro.Sheets("Spare1").Select()
                xlibro.Cells(3, 2) = "NYESTE" & hostNumberArray(3)

                xlibro.Sheets("Spare2").Select()
                xlibro.Cells(3, 2) = "NYESTE" & hostNumberArray(4)

                xlibro.Sheets("Primary").Select()

                strFilePathReg = GetSetting("DrawCheck", "UbiFilesDrawCheck", "").ToString()

                fileNameStr = dayNameStr & "_" & monthStr & dayStr & yearStr & "_" & "Check.xlsx"

                xlibro.ActiveWorkbook.SaveAs(strFilePathReg & "\" & fileNameStr)

                'xlibro.ActiveWorkbook.Save(strFilePathReg & "\" & fileNameStr)

                xlibro.Workbooks.Close()

                xlibro.Quit()


            Else

                fileNameStr = (strFilePathReg & "\" & fileNameStr)

                Select Case dateInt

                    Case 1

                        'fileNameStr = strPathFile & "\Sunday_SealCheck.xlsx"
                        xlibro.Workbooks.Open(fileNameStr)
                        xlibro.Visible = False
                        xlibro.Sheets("Primary").Select()
                        xlibro.Cells(1, 6) = fecha

                        fillExcelSundaySeal("Primary", xlibro)
                        fillExcelSundaySeal("Secondary", xlibro)
                        fillExcelSundaySeal("Spare", xlibro)
                        fillExcelSundaySeal("Spare1", xlibro)
                        fillExcelSundaySeal("Spare2", xlibro)


                    Case 2
                        'fileNameStr = strPathFile & "\ThursdayMonday_SealCheck.xlsx"
                        xlibro.Workbooks.Open(fileNameStr)
                        xlibro.Visible = False
                        xlibro.Sheets("Primary").Select()
                        xlibro.Cells(1, 6) = fecha


                        fillExcelThursdayMondaySeal("Primary", xlibro)
                        fillExcelThursdayMondaySeal("Secondary", xlibro)
                        fillExcelThursdayMondaySeal("Spare", xlibro)
                        fillExcelThursdayMondaySeal("Spare1", xlibro)
                        fillExcelThursdayMondaySeal("Spare2", xlibro)

                    Case 3

                        'fileNameStr = strPathFile & "\TuesdayFriday_SealCheck.xlsx"
                        xlibro.Workbooks.Open(fileNameStr)
                        xlibro.Visible = False
                        xlibro.Sheets("Primary").Select()
                        xlibro.Cells(1, 6) = fecha


                        fillExcelTuesdayFridaySeal("Primary", xlibro)
                        fillExcelTuesdayFridaySeal("Secondary", xlibro)
                        fillExcelTuesdayFridaySeal("Spare", xlibro)
                        fillExcelTuesdayFridaySeal("Spare1", xlibro)
                        fillExcelTuesdayFridaySeal("Spare2", xlibro)

                    Case 4

                        'fileNameStr = strPathFile & "\SaturdayWednesday_SealCheck.xlsx"
                        xlibro.Workbooks.Open(fileNameStr)
                        xlibro.Visible = False
                        xlibro.Sheets("Primary").Select()
                        xlibro.Cells(1, 6) = fecha

                        fillExcelSaturdayWednesdaySeal("Primary", xlibro)
                        fillExcelSaturdayWednesdaySeal("Secondary", xlibro)
                        fillExcelSaturdayWednesdaySeal("Spare", xlibro)
                        fillExcelSaturdayWednesdaySeal("Spare1", xlibro)
                        fillExcelSaturdayWednesdaySeal("Spare2", xlibro)


                    Case 5
                        'fileNameStr = strPathFile & "\ThursdayMonday_SealCheck.xlsx"
                        xlibro.Workbooks.Open(fileNameStr)
                        xlibro.Visible = False
                        xlibro.Sheets("Primary").Select()
                        xlibro.Cells(1, 6) = fecha

                        fillExcelThursdayMondaySeal("Primary", xlibro)
                        fillExcelThursdayMondaySeal("Secondary", xlibro)
                        fillExcelThursdayMondaySeal("Spare", xlibro)
                        fillExcelThursdayMondaySeal("Spare1", xlibro)
                        fillExcelThursdayMondaySeal("Spare2", xlibro)


                    Case 6

                        'fileNameStr = strPathFile & "\TuesdayFriday_SealCheck.xlsx"
                        xlibro.Workbooks.Open(fileNameStr)
                        xlibro.Visible = False
                        xlibro.Sheets("Primary").Select()
                        xlibro.Cells(1, 6) = fecha

                        fillExcelTuesdayFridaySeal("Primary", xlibro)
                        fillExcelTuesdayFridaySeal("Secondary", xlibro)
                        fillExcelTuesdayFridaySeal("Spare", xlibro)
                        fillExcelTuesdayFridaySeal("Spare1", xlibro)
                        fillExcelTuesdayFridaySeal("Spare2", xlibro)


                    Case 7

                        'fileNameStr = strPathFile & "\SaturdayWednesday_SealCheck.xlsx"
                        xlibro.Workbooks.Open(fileNameStr)
                        xlibro.Visible = False
                        xlibro.Sheets("Primary").Select()
                        xlibro.Cells(1, 6) = fecha


                        fillExcelSaturdayWednesdaySeal("Primary", xlibro)
                        fillExcelSaturdayWednesdaySeal("Secondary", xlibro)
                        fillExcelSaturdayWednesdaySeal("Spare", xlibro)
                        fillExcelSaturdayWednesdaySeal("Spare1", xlibro)
                        fillExcelSaturdayWednesdaySeal("Spare2", xlibro)

                End Select


                xlibro.Sheets("Primary").Select()
                xlibro.Cells(3, 2) = "NYESTE" & hostNumberArray(0)

                xlibro.Sheets("Secondary").Select()
                xlibro.Cells(3, 2) = "NYESTE" & hostNumberArray(1)

                xlibro.Sheets("Spare").Select()
                xlibro.Cells(3, 2) = "NYESTE" & hostNumberArray(2)

                xlibro.Sheets("Spare1").Select()
                xlibro.Cells(3, 2) = "NYESTE" & hostNumberArray(3)

                xlibro.Sheets("Spare2").Select()
                xlibro.Cells(3, 2) = "NYESTE" & hostNumberArray(4)

                xlibro.Sheets("Primary").Select()

                'strFilePathReg = GetSetting("DrawCheck", "UbiFilesDrawCheck", "").ToString()

                'fileNameStr = dayNameStr & "_" & monthStr & dayStr & yearStr & "_" & "Check.xlsx"

                'xlibro.ActiveWorkbook.SaveAs(strFilePathReg & "\" & fileNameStr)

                xlibro.ActiveWorkbook.Save()

                xlibro.Workbooks.Close()

                xlibro.Quit()


            End If

            Me.Cursor = Cursors.Default
            Me.RichTextBox1.Text = "All Done."
            Me.Refresh()


            fileNameStr = dayNameStr & "_" & monthStr & dayStr & yearStr & "_" & "Check.xlsx"
            strPathRemoteFile = "files/DrawCheck_Files"
            Me.RichTextBox1.Text = "Uploading to SFTP..." & vbCrLf
            Me.Refresh()


            errCodeInt = upLoadFileSFTP("10.1.5.182", "22", "xfer", "Welcome1", strPathRemoteFile & "/" & fileNameStr, strFilePathReg & "\" & fileNameStr, errMessageStr)


            Select Case errCodeInt

                Case 1
                    MsgBox("SFTP Component error. " + errMessageStr, MsgBoxStyle.OkOnly, "Error.")
                    Exit Sub
                Case 2

                    MsgBox("Conection error. " + errMessageStr, MsgBoxStyle.OkOnly, "Error.")
                    Exit Sub

                Case 3
                    MsgBox("Authenticate error. " + errMessageStr, MsgBoxStyle.OkOnly, "Error.")
                    Exit Sub

                Case 4
                    MsgBox("SFTP Initialize error. " + errMessageStr, MsgBoxStyle.OkOnly, "Error.")
                    Exit Sub


                Case 5
                    MsgBox("Upload file error. " + errMessageStr, MsgBoxStyle.OkOnly, "Error.")
                    Exit Sub

                Case 0

                    'file = My.Computer.FileSystem.OpenTextFileWriter(strFilePath & "LogTransferWinnnerAndJackpotFiles_" & monthStr & dayStr & yearStr & ".txt", True)
                    'file.WriteLine(Now & "  winner_summary_report_p010_d" & drawNumberStr & "_wincnt_english.rep" & " was successfully transferred")
                    'file.Close()

                    'Return 0
                    Me.Cursor = Cursors.Default
                    Me.RichTextBox1.Text = "All Done."
                    Me.Refresh()


            End Select


            Me.Dispose()


        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim errCodeInt As Integer = 0
        Dim errMessageStr As String = ""
        Dim file As System.IO.StreamWriter
        Dim infoStr As String = ""
        Dim xlibro As Microsoft.Office.Interop.Excel.Application
        Dim fecha As Date
        Dim dateInt As Integer
        Dim monthStr, dayStr, yearStr, fileNameStr, strPathFile, strFilePathReg, dayNameStr, strPathRemoteFile As String
        Dim hostNumberArray(5) As Integer


        fecha = Format(DateTimePicker1.Value, "MM/dd/yyyy")
        infoStr = getSealFromSSHNew("10.1.5.11", "cvega", "Bogot@1234", fecha, errCodeInt, errMessageStr)
        If errCodeInt <> 0 Then
            MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        End If

        file = My.Computer.FileSystem.OpenTextFileWriter("SEAL_PrimaryTest.txt", False)
        file.WriteLine(infoStr)
        file.Close()

        MsgBox(infoStr)





        'getHostModeNew("10.1.5.11", "cvega", "Bogot@1234", hostNumberArray, errCodeInt, errMessageStr)

        'If errCodeInt < 0 Then
        '    Select Case errCodeInt
        '        Case -1

        '            MsgBox(errMessageStr)


        '        Case -2

        '            MsgBox(errMessageStr)


        '        Case -3

        '            MsgBox(errMessageStr)


        '        Case -4

        '            MsgBox(errMessageStr)

        '        Case -5

        '            MsgBox(errMessageStr)


        '        Case -6

        '            MsgBox(errMessageStr)


        '        Case -7

        '            MsgBox(errMessageStr)


        '        Case -8

        '            MsgBox(errMessageStr)


        '    End Select

        'Else

        '    infoStr = getSealFromSSHNew("10.1.5.11", "cvega", "Bogot@1234", fecha, errCodeInt, errMessageStr)
        '    If errCodeInt <> 0 Then
        '        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
        '    End If

        '    MsgBox(infoStr)

        'End If



    End Sub
End Class
