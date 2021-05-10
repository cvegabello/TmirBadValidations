Imports System
Imports System.Reflection

Public Class frmMain

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




    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim cdcInt As Integer

        'Shell("NET USE H: \\10.5.165.98\SharedDocuments Formula7 /USER:gtk\cvegabello")

        DateTimePicker1.Value = Now.AddDays(-1) 'Normalmente debe estar con -1, es decir con la fecha anterior

        cdcInt = returnCDC(DateTimePicker1.Value)
        Label1.Text = cdcInt
        GroupBox1.Enabled = False
        ProgressBar1.Maximum = 100
        okBtn.Visible = False
        cancelBtn.Visible = False
        Timer1.Enabled = True



    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Dim errCode As Integer
        Dim errorStr As String = ""
        Dim cdcStr As String
        Dim hostNumberArray(5) As Integer
        Dim conStrinHostStr As String
        Dim substrings() As String



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

        Static count As Integer = 0
        count = count + 10
        If count <= 100 Then
            ProgressBar1.Value = count
        Else
            Timer1.Enabled = False

            cdcStr = Trim(Label1.Text)

            getHostModeNew(substrings(3), substrings(0), substrings(4), hostNumberArray, errCode, errorStr)

            'Utils.getHostMode("10.1.5.12", "prosys", "Numb3r1j0b", hostNumberArray, errCode, errorStr)

            'Dim result As Integer = MessageBox.Show("Funciono getHostMode el spare PDC system is: " & hostNumberArray(2), "Funciono", MessageBoxButtons.OK)
            'If result = DialogResult.OK Then
            '    Me.Dispose()

            'End If


            If errCode = 0 Then
                'MsgBox(hostNumberArray(2))

                Select Case hostNumberArray(2)

                    Case 1
                        Utils.auto_IVAL_IPS_gxtmir_New(substrings(1), substrings(0), substrings(2), cdcStr, errCode, errorStr)

                    Case 2

                        Utils.auto_IVAL_IPS_gxtmir_New(substrings(3), substrings(0), substrings(4), cdcStr, errCode, errorStr)

                    Case 3
                        Utils.auto_IVAL_IPS_gxtmir_New(substrings(5), substrings(0), substrings(6), cdcStr, errCode, errorStr)

                    Case 4
                        Utils.auto_IVAL_IPS_gxtmir_New(substrings(7), substrings(0), substrings(8), cdcStr, errCode, errorStr)

                    Case 5
                        Utils.auto_IVAL_IPS_gxtmir_New(substrings(9), substrings(0), substrings(10), cdcStr, errCode, errorStr)


                End Select
                'If errCode <> 0 Then
                '    MsgBox("Error: " & errorStr)
                'End If

            Else

                MsgBox(errorStr, MsgBoxStyle.Critical, "GetHostMode Error")


            End If

            Me.Dispose()

        End If

    End Sub

    Private Sub okBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles okBtn.Click

        Dim errCode As Integer
        Dim errorStr As String = ""
        Dim cdcStr As String
        Dim hostNumberArray(5) As Integer
        Dim conStrinHostStr As String
        Dim substrings() As String

        conStrinHostStr = GetSettingConfigHost(appName, "conStringHost", "").ToString()
        substrings = conStrinHostStr.Split("|")

        cdcStr = Trim(Label1.Text)

        getHostModeNew(substrings(3), substrings(0), substrings(4), hostNumberArray, errCode, errorStr)
        'Utils.getHostMode("10.1.5.12", "prosys", "Numb3r1j0b", hostNumberArray, errCode, errorStr)


        If errCode = 0 Then
            'MsgBox(hostNumberArray(2))

            Select Case hostNumberArray(2)

                Case 1
                    Utils.auto_IVAL_IPS_gxtmir_New(substrings(1), substrings(0), substrings(2), cdcStr, errCode, errorStr)
                    'Utils.auto_IVAL_IPS_gxtmir("10.1.5.10", "prosys", "Numb3r1j0b", cdcStr, errCode, errorStr)
                    If errCode <> 0 Then
                        MsgBox("Error: " & errorStr)
                    End If

                Case 2

                    Utils.auto_IVAL_IPS_gxtmir_New(substrings(3), substrings(0), substrings(4), cdcStr, errCode, errorStr)
                    'Utils.auto_IVAL_IPS_gxtmir("10.1.5.11", "prosys", "Numb3r1j0b", cdcStr, errCode, errorStr)
                    If errCode <> 0 Then
                        MsgBox("Error: " & errorStr)
                    End If


                Case 3
                    Utils.auto_IVAL_IPS_gxtmir(substrings(5), substrings(0), substrings(6), cdcStr, errCode, errorStr)
                    If errCode <> 0 Then
                        MsgBox("Error: " & errorStr)
                    End If


                Case 4
                    Utils.auto_IVAL_IPS_gxtmir(substrings(7), substrings(0), substrings(8), cdcStr, errCode, errorStr)
                    If errCode <> 0 Then
                        MsgBox("Error: " & errorStr)
                    End If


                Case 5
                    Utils.auto_IVAL_IPS_gxtmir(substrings(9), substrings(0), substrings(10), cdcStr, errCode, errorStr)
                    If errCode <> 0 Then
                        MsgBox("Error: " & errorStr)
                    End If



            End Select

        Else

            MsgBox(errorStr, MsgBoxStyle.Critical, "GetHostMode Error")


        End If


    End Sub

    Private Sub stopBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles stopBtn.Click
        Timer1.Enabled = False
        stopBtn.Visible = False
        ProgressBar1.Visible = False
        GroupBox1.Enabled = True

        okBtn.Visible = True
        cancelBtn.Visible = True
    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        Dim dateSelected As Date
        Dim cdcInt As Integer


        dateSelected = DateTimePicker1.Value
        cdcInt = returnCDC(dateSelected)
        Label1.Text = cdcInt

    End Sub

    Private Sub cancelBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cancelBtn.Click
        Me.Dispose()
    End Sub

    
End Class
