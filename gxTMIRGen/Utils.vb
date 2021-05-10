Imports System.IO
Imports System.Text
Imports System.Xml
Imports Microsoft.Win32
Imports System.Net.Mail


Module Utils
    Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
    Public appName As String = "SETTINGS LOCAL APPS"

    Public Function GetSetting(ByVal APP_NAME As String, ByVal Keyname As String, Optional ByVal DefVal As String = "") As String
        Dim Key As RegistryKey
        Try
            Key = Registry.CurrentUser.CreateSubKey("SOFTWARE").CreateSubKey("IGT").CreateSubKey(APP_NAME)
            Return Key.GetValue(Keyname, DefVal)


        Catch
            Return DefVal
        End Try
    End Function

    Public Sub SalvarSetting(ByVal APP_NAME As String, ByVal Keyname As String, ByVal Value As String)
        Dim Key As RegistryKey
        Try
            Key = Registry.CurrentUser.CreateSubKey("SOFTWARE").CreateSubKey("IGT").CreateSubKey(APP_NAME)
            Key.SetValue(Keyname, Value)
        Catch
            Return
        End Try
    End Sub

    Public Sub SalvarSettingConfiServerDB(ByVal APP_NAME As String, ByVal Keyname As String, ByVal Value As String)
        Dim Key As RegistryKey
        Try
            Key = Registry.CurrentUser.CreateSubKey("SOFTWARE").CreateSubKey("IGT").CreateSubKey(APP_NAME).CreateSubKey("Conection string database")
            Key.SetValue(Keyname, Value)
        Catch
            Return
        End Try
    End Sub

    Public Function GetSettingConfigServerDB(ByVal APP_NAME As String, ByVal Keyname As String, Optional ByVal DefVal As String = "") As String
        Dim Key As RegistryKey
        Try
            Key = Registry.CurrentUser.CreateSubKey("SOFTWARE").CreateSubKey("IGT").CreateSubKey(APP_NAME).CreateSubKey("Conection string database")
            Return Key.GetValue(Keyname, DefVal)
        Catch
            Return DefVal
        End Try
    End Function

    Public Sub SalvarSettingConfiHost(ByVal APP_NAME As String, ByVal Keyname As String, ByVal Value As String)
        Dim Key As RegistryKey
        Try
            Key = Registry.CurrentUser.CreateSubKey("SOFTWARE").CreateSubKey("IGT").CreateSubKey(APP_NAME).CreateSubKey("Conection string host")
            Key.SetValue(Keyname, Value)
        Catch
            Return
        End Try
    End Sub

    Public Function GetSettingConfigHost(ByVal APP_NAME As String, ByVal Keyname As String, Optional ByVal DefVal As String = "") As String
        Dim Key As RegistryKey
        Try
            Key = Registry.CurrentUser.CreateSubKey("SOFTWARE").CreateSubKey("IGT").CreateSubKey(APP_NAME).CreateSubKey("Conection string host")
            Return Key.GetValue(Keyname, DefVal)
        Catch
            Return DefVal
        End Try
    End Function

    Sub auto_IVAL_IPS_gxtmir_New(ByVal hostName As String, ByVal userName As String, ByVal password As String, ByVal cdcStr As String, ByRef errCodeInt As Integer, ByRef errMessageStr As String)


        Dim ssh As New Chilkat.Ssh()
        Dim port As Integer
        Dim channelNum As Integer
        Dim termType As String = "vt100"
        Dim widthInChars As Integer = 80
        Dim heightInChars As Integer = 24
        Dim pixWidth As Integer = 0
        Dim pixHeight As Integer = 0
        Dim cmdOutputStr As String = ""
        Dim msgError As String = ""
        Dim n As Integer
        Dim pollTimeoutMs As Integer = 8000
        Dim strPathRemoteFile, strPathFile As String


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


        errCodeInt = SSHUntilMatch(ssh, "Enter Selection:", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Sub
            'Else
            '    MsgBox("FUNCIONO: ", MsgBoxStyle.OkOnly, "FUNCIONO.")

        End If

        errCodeInt = sentStringToSSHUntilMatch(ssh, "98", "return to the menu):", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Sub
            'Else
            '    MsgBox("FUNCIONO: 10 ", MsgBoxStyle.OkOnly, "FUNCIONO.")
        End If


        errCodeInt = sentStringToSSHUntilMatch(ssh, "sudo su - prosys", "Password:", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Sub
            'Else
            '    MsgBox("FUNCIONO: 01 ", MsgBoxStyle.OkOnly, "FUNCIONO.")


        End If

        errCodeInt = sentStringToSSH(ssh, password, channelNum, pollTimeoutMs, cmdOutputStr, msgError)

        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Sub
        End If




        '=================================

        errCodeInt = sentStringToSSH(ssh, "cd ival/bin", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Sub

        End If
        Sleep(1500)


        errCodeInt = sentStringToSSH(ssh, "ival_ips_gxtmir", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Sub

        End If

        errCodeInt = sentStringToSSHUntilMatch(ssh, cdcStr, ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "y", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "3", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", "*End of task*", channelNum, pollTimeoutMs, cmdOutputStr, msgError)

        'errCodeInt = 0

        If errCodeInt = 0 Then

            errCodeInt = sentStringToSSH(ssh, "cd oper/bin", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
            If errCodeInt <> 0 Then
                errCodeInt = -9
                errMessageStr = ssh.LastErrorText
                Exit Sub

            End If
            Sleep(1000)

            errCodeInt = sentStringToSSHUntilMatch(ssh, "ny_esa_inquiry.ksh", "oltp/oper/bin", channelNum, pollTimeoutMs, cmdOutputStr, msgError)

            If errCodeInt <> 0 Then
                errCodeInt = -9
                errMessageStr = ssh.LastErrorText
                Exit Sub

            End If

            'strPathRemoteFile = "/rptfiles1/proips/ival/reports/" & "ival_ips_gxtmir_" & cdcStr & ".rep"
            'strPathFile = "F:\ival_ips_gxtmir_" & cdcStr & ".rep"
            'errCodeInt = downloadFileSFTP(hostName, "22", userName, password, strPathRemoteFile, strPathFile, errMessageStr)
            ''frmMain.TextBox1.Text = "Error Code: " & errCodeInt & " " & errMessageStr
            ''MsgBox("Error Code: " & errCodeInt & " " & errMessageStr)
            'If errCodeInt <> 0 Then
            '    errCodeInt = -10
            '    errMessageStr = ssh.LastErrorText
            '    Exit Sub
            'End If
        Else
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Sub

        End If

        ssh.Disconnect()

    End Sub




    Sub auto_IVAL_IPS_gxtmir(ByVal hostName As String, ByVal userName As String, ByVal password As String, ByVal cdcStr As String, ByRef errCodeInt As Integer, ByRef errMessageStr As String)


        Dim ssh As New Chilkat.Ssh()
        Dim port As Integer
        Dim channelNum As Integer
        Dim termType As String = "vt100"
        Dim widthInChars As Integer = 80
        Dim heightInChars As Integer = 24
        Dim pixWidth As Integer = 0
        Dim pixHeight As Integer = 0
        Dim cmdOutputStr As String = ""
        Dim msgError As String = ""
        Dim n As Integer
        Dim pollTimeoutMs As Integer = 8000
        Dim strPathRemoteFile, strPathFile As String


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
            Exit Sub

        End If


        errCodeInt = sentStringToSSH(ssh, "cd ival/bin", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Sub

        End If
        Sleep(1500)


        errCodeInt = sentStringToSSH(ssh, "ival_ips_gxtmir", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Sub

        End If

        errCodeInt = sentStringToSSHUntilMatch(ssh, cdcStr, ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "y", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "3", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", ">", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        Sleep(1000)
        'MsgBox(cmdOutputStr)
        errCodeInt = sentStringToSSHUntilMatch(ssh, "", "*End of task*", channelNum, pollTimeoutMs, cmdOutputStr, msgError)


        If errCodeInt = 0 Then

            errCodeInt = sentStringToSSH(ssh, "cd ival/reports", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
            If errCodeInt <> 0 Then
                errCodeInt = -9
                errMessageStr = ssh.LastErrorText
                Exit Sub

            End If
            Sleep(1000)
            errCodeInt = sentStringToSSHUntilMatch(ssh, "tar -cvzf ival_ips_gxtmir_" & cdcStr & ".rep.tgz ival_ips_gxtmir_" & cdcStr & ".rep", "*/reports*", channelNum, pollTimeoutMs, cmdOutputStr, msgError)

            'errCodeInt = 0 ' No olvidar comentariar

            If errCodeInt = 0 Then
                'MsgBox("TAR DONE")

                strPathRemoteFile = "/rptfiles1/proips/ival/reports/" & "ival_ips_gxtmir_" & cdcStr & ".rep.tgz"

                'strPathFile = Trim(ubiWinFilesTxt.Text)
                strPathFile = "W:\ival_ips_gxtmir_" & cdcStr & ".rep.tgz"
                errCodeInt = downloadFileSFTP(hostName, "22", userName, password, strPathRemoteFile, strPathFile, errMessageStr)
                If errCodeInt <> 0 Then
                    errCodeInt = -10
                    errMessageStr = ssh.LastErrorText
                    Exit Sub

                End If

                errCodeInt = sentStringToSSH(ssh, "cd oper/bin", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
                If errCodeInt <> 0 Then
                    errCodeInt = -9
                    errMessageStr = ssh.LastErrorText
                    Exit Sub

                End If
                Sleep(1000)

                errCodeInt = sentStringToSSHUntilMatch(ssh, "ny_esa_inquiry.ksh", "oltp/oper/bin", channelNum, pollTimeoutMs, cmdOutputStr, msgError)

                If errCodeInt <> 0 Then
                    errCodeInt = -9
                    errMessageStr = ssh.LastErrorText
                    Exit Sub

                End If

            Else
                errCodeInt = -9
                errMessageStr = ssh.LastErrorText
                Exit Sub

            End If


        Else
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Sub

        End If


        ssh.Disconnect()

    End Sub


    Sub getHostMode(ByVal hostName As String, ByVal userName As String, ByVal password As String, ByVal hostNumberArray() As Integer, ByRef errCodeInt As Integer, ByRef errMessageStr As String)


        Dim ssh As New Chilkat.Ssh
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



  


    Function getDrawNumberXProductXday(ByVal codSysProduct As Integer, ByVal toKnowDate As DateTime, ByVal dayDraw As Integer) As Integer

        Dim num1 As Integer
        Dim num3 As Integer
        Dim dayWeekInt As Integer
        Dim moduloRes As Integer
        Dim cocienteRes As Integer

        num1 = convertDateInt(toKnowDate)
        dayWeekInt = Weekday(toKnowDate)

        Select Case codSysProduct
            Case 8 'Loto

                Select Case dayWeekInt
                    Case 5
                        moduloRes = (num1 + 2) Mod 7
                        cocienteRes = ((num1 + 2) \ 7) * 2
                    Case 6
                        moduloRes = (num1 + 1) Mod 7
                        cocienteRes = ((num1 + 1) \ 7) * 2
                    Case Else
                        moduloRes = num1 Mod 7
                        cocienteRes = (num1 \ 7) * 2

                End Select
                If (moduloRes < 1) Then
                    num3 = cocienteRes - 11488
                Else
                    num3 = cocienteRes - 11487
                End If
                num3 += 2357



            Case 9 'PCK3 -NUMBERS-

                Select Case dayDraw
                    Case 1 'Midday
                        num3 = ((2 * num1) - 69216)
                    Case 2 'Evenning
                        num3 = (((2 * num1) - 69216) + 1)
                End Select

            Case 10
                num3 = (num1 - 35525)

            Case 11

            Case 12 'BIGG
                If ((num1 Mod 7) < 4) Then
                    num3 = (((num1 \ 7) * 2) - 10682)
                Else
                    num3 = (((num1 \ 7) * 2) - 10681)
                End If

            Case 13 'LIFE

                Select Case dayWeekInt
                    Case 6
                        num3 = (((num1 \ 7) * 2) - 11941)

                    Case Else
                        If ((num1 Mod 7) < 3) Then
                            num3 = (((num1 \ 7) * 2) - 11943)
                        Else
                            num3 = (((num1 \ 7) * 2) - 11942)
                        End If

                End Select


            Case 14

                Select Case dayDraw
                    Case 1 'Midday
                        num3 = ((2 * num1) - 69216)
                    Case 2 'Evenning
                        num3 = (((2 * num1) - 69216) + 1)
                End Select


            Case 15 'PowerBall

                Select Case dayWeekInt
                    Case 5
                        moduloRes = (num1 + 2) Mod 7
                        cocienteRes = ((num1 + 2) \ 7) * 2
                    Case 6
                        moduloRes = (num1 + 1) Mod 7
                        cocienteRes = ((num1 + 1) \ 7) * 2
                    Case Else
                        moduloRes = num1 Mod 7
                        cocienteRes = (num1 \ 7) * 2

                End Select
                If (moduloRes < 1) Then
                    num3 = cocienteRes - 11488
                Else
                    num3 = cocienteRes - 11487
                End If

            Case 27 'DKNO
                num3 = (num1 - 31979)

        End Select
        Return num3
    End Function


    Function convertDateInt(ByVal toKnowDate As Date) As Integer
        Dim initDate As Date
        Dim diff1 As TimeSpan
        initDate = "1/1/1900"

        diff1 = toKnowDate.Subtract(initDate)
        convertDateInt = diff1.Days + 2

    End Function
    Function returnCDC(ByVal toKnowDate As Date) As Integer
        Dim initDate As Date
        Dim diff1 As TimeSpan
        initDate = "1/1/1986"

        diff1 = toKnowDate.Subtract(initDate)
        returnCDC = diff1.Days + 1

    End Function

    Function sent_Email(ByVal smtpHostStr As String, ByVal toStr As String, ByVal attachStr As String, ByVal subjectStr As String, ByVal bodyStr As String) As Integer
        Dim SMTPserver As New SmtpClient
        Dim mail As New MailMessage
        Dim oAttch As Attachment = New Attachment(attachStr)


        Try
            SMTPserver.Host = smtpHostStr
            mail = New MailMessage
            mail.From = New MailAddress("do.not.reply@gtech-noreply.com")
            mail.To.Add(toStr) 'The Man you want to send the message to him
            mail.Subject = subjectStr
            mail.Body = bodyStr
            mail.Attachments.Add(oAttch)
            SMTPserver.Send(mail)
            Return 0
            'MessageBox.Show("Done!", "Message Sent", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)


        Catch ex As Exception
            Return -1

        End Try




    End Function

End Module
