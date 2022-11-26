Public Class frmConfigScrn
    '***********  Communication Settings Configuration Form

    Private NewData As Int32
    Private NewStop As Int32
    Private NewParity As Int32
    '
    '--- No parity option button
    '
    Private Sub NoParity_CheckedChanged(ByVal eventSender _
    As System.Object, ByVal eventArgs As System.EventArgs) _
                            Handles NoParity.CheckedChanged
        If eventSender.Checked Then
            NewParity = IO.Ports.Parity.None
        End If
    End Sub
    '
    '--- Odd parity option button
    '
    Private Sub OddParity_CheckedChanged(ByVal eventSender _
    As System.Object, ByVal eventArgs As System.EventArgs) _
                            Handles OddParity.CheckedChanged
        If eventSender.Checked Then
            NewParity = IO.Ports.Parity.Odd
        End If
    End Sub
    '
    '--- Even parity option button
    '
    Private Sub EvenParity_CheckedChanged(ByVal eventSender _
    As System.Object, ByVal eventArgs As System.EventArgs) _
                            Handles EvenParity.CheckedChanged
        If eventSender.Checked Then
            NewParity = IO.Ports.Parity.Even
        End If
    End Sub

    '--- Initialize and display configuration form
    '
    Private Sub frmConfigScrn_Load(ByVal eventSender As _
    System.Object, ByVal eventArgs As System.EventArgs) _
                                    Handles MyBase.Load
        Dim I As Short
        'For I = 1 To 255
        '    lstCommPort.Items.Add("COM" & I.ToString)
        'Next I
        Dim PortNames() As String = IO.Ports.SerialPort.GetPortNames()
        For I = PortNames.Length - 1 To 0 Step -1
            lstCommPort.Items.Add(PortNames(I))
        Next
        lstRate.Items.Add("300")
        lstRate.Items.Add("1200")
        lstRate.Items.Add("2400")
        lstRate.Items.Add("4800")
        lstRate.Items.Add("9600")
        lstRate.Items.Add("19200")
        lstRate.Items.Add("38400")
        lstRate.Items.Add("57600")
        lstRate.Items.Add("115200")

        With NETSerialTerm.SerialPort '*Pads.SerialPort 'NETSerialTerm.SerialPort
            '--- Get current port
            lstCommPort.SelectedIndex = _
                    lstCommPort.FindString(.PortName)

            '--- Get current rate
            Select Case .BaudRate 'select rate
                Case 300 'set active baud
                    lstRate.SelectedIndex = 0
                Case 1200
                    lstRate.SelectedIndex = 1
                Case 2400
                    lstRate.SelectedIndex = 2
                Case 4800
                    lstRate.SelectedIndex = 3
                Case 9600
                    lstRate.SelectedIndex = 4
                Case 19200
                    lstRate.SelectedIndex = 5
                Case 38400
                    lstRate.SelectedIndex = 6
                Case 57600
                    lstRate.SelectedIndex = 7
                Case 115200
                    lstRate.SelectedIndex = 8
                Case Else
                    lstRate.SelectedIndex = 4
            End Select

            '            IO.Ports.Parity.None()

            '--- Get current parity
            NewParity = .Parity
            Select Case .Parity
                Case .Parity.None   'set active parity
                    NoParity.Checked = True 'option button
                Case .Parity.Even
                    EvenParity.Checked = True
                Case .Parity.Odd
                    OddParity.Checked = True
            End Select

            '--- Get data bits
            NewData = .DataBits
            Select Case .DataBits 'select data bits
                Case 7 'set active choice
                    Data7.Checked = True 'option button
                Case 8
                    Data8.Checked = True
            End Select

            '--- Get stop bits
            NewStop = .StopBits
            Select Case .StopBits  'select stop bits
                Case .StopBits.One 'set active choice
                    Stop1.Checked = True 'option button
                Case .StopBits.Two
                    Stop2.Checked = True
            End Select
        End With
    End Sub
    '
    '--- 1 stop bit option button
    '
    Private Sub Stop1_CheckedChanged(ByVal eventSender _
    As System.Object, ByVal eventArgs As System.EventArgs) _
                            Handles Stop1.CheckedChanged
        If eventSender.Checked Then
            NewStop = IO.Ports.StopBits.One
        End If
    End Sub
    '
    '--- 2 stop bits option button
    '
    Private Sub Stop2_CheckedChanged(ByVal eventSender As _
    System.Object, ByVal eventArgs As System.EventArgs) _
                            Handles Stop2.CheckedChanged
        If eventSender.Checked Then
            NewStop = IO.Ports.StopBits.Two
        End If
    End Sub
    '
    '--- 8 data bits option button
    '
    Private Sub Data8_CheckedChanged(ByVal sender As _
        System.Object, ByVal e As System.EventArgs) _
        Handles Data8.CheckedChanged
        NewData = 8
    End Sub
    '
    '--- 7 data bits option button
    '
    Private Sub Data7_CheckedChanged(ByVal sender As _
        System.Object, ByVal e As System.EventArgs) _
        Handles Data7.CheckedChanged
        NewData = 7
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e _
    As System.EventArgs) Handles Button1.Click
        '
        '--- Ok button actions
        '
        Dim OldPort As String
        Dim PortOpen As Boolean

        With Pads.SerialPort '#NETSerialTerm.SerialPort
            OldPort = .PortName
            PortOpen = .IsOpen
            If PortOpen = True Then .Close()
            .BaudRate = Val(lstRate.Text)
            .Parity = NewParity
            .DataBits = NewData
            .StopBits = NewStop
            .PortName = lstCommPort.SelectedItem
            'set new port number
            If PortOpen = True Then
                Try
                    .Open()
                Catch Ex As Exception
                    MsgBox(Err.Description)
                Finally
                    If .IsOpen = False Then
                        MsgBox("Selected port could not be opened", MsgBoxStyle.Exclamation)
                        .PortName = OldPort
                    End If
                End Try
            End If
        End With
        Me.Close() 'remove configuration form
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e _
    As System.EventArgs) Handles Button2.Click
        '--- Cancel button actions
        '
        '
        Me.Close()
    End Sub
End Class