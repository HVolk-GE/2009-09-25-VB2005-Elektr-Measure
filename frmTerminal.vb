Imports System.IO.Ports
Public Class NETSerialTerm
    Public Shared WithEvents SerialPort As SerialPort
    Private Shared m_FormDefInstance As NETSerialTerm
    Private Shared m_InitializingDefInstance As Boolean

    Public Shared Property DefInstance() As NETSerialTerm
        Get
            If m_FormDefInstance Is Nothing OrElse _
                        m_FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_FormDefInstance = New NETSerialTerm
                m_InitializingDefInstance = False
            End If
            DefInstance = m_FormDefInstance
        End Get
        Set(ByVal Value As NETSerialTerm)
            m_FormDefInstance = Value
        End Set
    End Property

    Private Sub PortOpenToolStripMenuItem_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles PortOpenToolStripMenuItem.Click
        Dim ex As Exception
        With SerialPort
            If .IsOpen = False Then
                Try
                    .Open()
                Catch ex
                End Try
            Else
                Try
                    .Close()
                Catch ex
                End Try
            End If
            If .IsOpen = True Then
                PortOpenToolStripMenuItem.Checked = True
                Me.Text = "NETSerialTerm using port: " & _
                                    SerialPort.PortName
                .RtsEnable = True
                .DtrEnable = True
                .ReceivedBytesThreshold = 1
            Else
                PortOpenToolStripMenuItem.Checked = False
                Me.Text = "NETSerialTerm not running"
            End If
        End With
    End Sub

    Private Sub txtTerm_KeyPress(ByVal sender As Object, ByVal e As _
    System.Windows.Forms.KeyPressEventArgs) Handles txtTerm.KeyPress
        Dim KeyAscii As Int32 = Asc(e.KeyChar)
        With SerialPort
            If .IsOpen = True Then
                .Write(Chr(KeyAscii))
            End If
        End With
        e.Handled = True
    End Sub


    Public Delegate Sub DisplayData(ByVal Buffer As String)

    Private Shared Sub Display(ByVal Buffer As String)
        Buffer = Buffer.Replace(vbLf, vbCr)
        Buffer = Buffer.Replace(vbCr & vbCr, vbCr)
        Buffer = Buffer.Replace(vbCr, vbCrLf)
        With DefInstance.txtTerm
            If (Buffer.Length = 1) And (Buffer = Chr(8)) _
                                                    Then
                If (.Text.Length > 0) Then .Text = _
                    .Text.Remove(.Text.Length - 1, 1)
            Else
                .AppendText(Buffer)
            End If
            If .Text.Length > 8196 Then
                .Text = .Text.Remove(0, 4096)
                If Mid(.Text, 1) = vbLf Then _
                        .Text = .Text.Remove(0, 1)
            End If
            .SelectionStart = .Text.Length
        End With
    End Sub

    Private Sub NETSerialTerm_FormClosing(ByVal sender As Object, ByVal e As _
    System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If SerialPort.IsOpen Then SerialPort.Close()
    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SerialPort = New SerialPort
        DefInstance = Me
    End Sub

    Private Shared Sub SerialPort_DataReceived(ByVal sender As Object, ByVal e As _
    System.IO.Ports.SerialDataReceivedEventArgs) Handles SerialPort.DataReceived
        Dim Buffer As String = SerialPort.ReadExisting()
        DefInstance.txtTerm.BeginInvoke(New _
            DisplayData(AddressOf Display), _
                    New Object() {Buffer})
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As _
    System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
        MainMenue.Show()
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As _
    System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        MsgBox("NETSerialTerm is a simple terminal emulator that illustrates" & _
        vbCrLf & "the Visual Studio 2005 System.IO.Ports serial IO class." & vbCrLf _
        & "Copyright (c) 2005 by Mabry Software, Inc.")
    End Sub

    Private Sub SettingsToolStripMenuItem_Click(ByVal sender As _
    System.Object, ByVal e As System.EventArgs) Handles SettingsToolStripMenuItem.Click
        frmConfigScrn.ShowDialog()
        Me.Text = "NETSerialTerm using port: " & _
                    SerialPort.PortName
    End Sub

    Private Sub ClearScreeToolStripMenuItem_Click(ByVal sender As _
    System.Object, ByVal e As System.EventArgs) Handles ClearScreeToolStripMenuItem.Click
        txtTerm.Text = ""
    End Sub
End Class
