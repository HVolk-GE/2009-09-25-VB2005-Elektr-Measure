Option Strict Off
Option Explicit On
Imports System.IO.Ports

Friend Class Pads
    Inherits System.Windows.Forms.Form
    Public Shared WithEvents SerialPort As SerialPort
    Private Shared m_FormDefInstance As Pads
    Private Shared m_InitializingDefInstance As Boolean
    Public Measurepoint As Integer

    Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
        canchelPad = 0
        Me.Close()
        Form2.Show()
    End Sub

    Public Sub Pads_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        If SerialPort.IsOpen Then SerialPort.Close()

    End Sub

    Private Shared Sub SerialPort_DataReceived(ByVal sender As Object, ByVal e As _
       System.IO.Ports.SerialDataReceivedEventArgs) Handles SerialPort.DataReceived
        Dim Buffer As String = SerialPort.ReadExisting()
        DefInstance.Text1.BeginInvoke(New _
            DisplayData(AddressOf Display), _
                    New Object() {Buffer})
    End Sub

    Public Shared Property DefInstance() As Pads
        Get
            If m_FormDefInstance Is Nothing OrElse _
                        m_FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_FormDefInstance = New Pads
                m_InitializingDefInstance = False
            End If
            DefInstance = m_FormDefInstance
        End Get
        Set(ByVal Value As Pads)
            m_FormDefInstance = Value
        End Set
    End Property

    Private Sub Pads_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Measurepoint = 1
        SerialPort = New SerialPort
        DefInstance = Me
        Me.ComboBox1.Text = ""
        Me.ComboBox2.Text = ""

        tempstr0 = ""

        Me.Check1.BackColor = Color.Green
        Me.CheckBox1.BackColor = Color.Green

        If State = "DE" Then
            Me.Label3.Text = "Kontrolle :"
            Me.Check1.Text = "Start Messung"
            Me.Label1.Text = "Benutzername"
            Me.Label2.Text = "Gewicht (g)"
            Me.Command1.Text = "Schliessen"
            Me.Button1.Text = "OK"
            Me.Check1.Text = "Start Messung"
            Me.Label4.Text = "Messinst. ID"
            Me.Label5.Text = "Waage ID"
            Me.CheckBox1.Text = "Start Wiegen"
            msgMeasurend = "Anzahl der Messpunkte erreicht, Messung beendet !"
        Else
            Me.Label3.Text = "Control :"
            Me.Label1.Text = "Username"
            Me.Label2.Text = "Weight (g)"
            Me.Command1.Text = "Close"
            Me.Button1.Text = "OK"
            Me.Check1.Text = "Start Measurement"
            Me.Label4.Text = "Meas.Inst. ID"
            Me.Label5.Text = "Weigh ID"
            Me.CheckBox1.Text = "Start Weigh"
            msgMeasurend = "Measruement finished !"
        End If

        Me.Text19.Text = strUsername
        'Me.Text1.Focus = True
        Me.Text1.SelectAll()
        Me.Text1.BackColor = Color.LightSeaGreen
        Me.Text2.BackColor = Color.LightSeaGreen
        Me.Text3.BackColor = Color.LightSeaGreen
        Me.Text4.BackColor = Color.LightSeaGreen
        Me.Text5.BackColor = Color.LightSeaGreen
        Me.Text6.BackColor = Color.LightSeaGreen
        Me.Text7.BackColor = Color.LightSeaGreen
        Me.Text8.BackColor = Color.LightSeaGreen
        Me.Text9.BackColor = Color.LightSeaGreen
        Me.Text10.BackColor = Color.LightSeaGreen
        Me.Text11.BackColor = Color.LightSeaGreen
        Me.Text12.BackColor = Color.LightSeaGreen
        Me.Text13.BackColor = Color.LightSeaGreen
        Me.TextBox4.BackColor = Color.LightSeaGreen
        Me.ComboBox1.BackColor = Color.LightSeaGreen
        Me.ComboBox2.BackColor = Color.LightSeaGreen
        'Me.TextBox1.Width = Len(Me.TextBox1.Text) + 4

        'Me.Option1.Checked = True
        'Me.Option1.BackColor = Color.DarkRed

        Me.TextBox2.BackColor = Color.LightSeaGreen
        Me.TextBox2.Text = ""
        Me.TextBox2.Enabled = True

        Me.Label7.Visible = False
        Me.Label8.Visible = False
        Me.Label9.Visible = False

        Me.ComboBox1.Items.Add(MitutoyoInstr1ID)
        Me.ComboBox1.Items.Add(MitutoyoInstr2ID)
        Me.ComboBox1.Items.Add(MitutoyoInstr3ID)
        Me.ComboBox1.Items.Add(MitutoyoInstr4ID)

        Me.ComboBox2.Items.Add(weightID)
        Me.ComboBox2.Items.Add(weight1ID)

        lngNumCols = CInt(Form2.Label2.Text)

        Me.TopMost = True

        Exit Sub

err_Handler:
        'Wenn ein Fehler aufgetreten ist, alle Controls
        '  ausser Beenden-Button deaktivieren:
        Dim ctl As System.Windows.Forms.Control

        For Each ctl In Me.Controls
            If ctl.Name <> "Command1" Then
                ctl.Enabled = False
            End If
        Next ctl
        Exit Sub
    End Sub

    Private Sub CheckBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox1.Click
        Dim ex As Exception

        If Me.Check1.Checked = True Then
            Me.Check1.BackColor = Color.Green
            Me.Check1.Checked = False

            If SerialPort.IsOpen Then SerialPort.Close()

            If State = "DE" Then
                Me.CheckBox1.Text = "Start Wiegen"
                Me.Check1.Text = "Start Messung"
            Else
                Me.CheckBox1.Text = "Start Weight"
                Me.Check1.Text = "Start Measurement"
            End If

        End If

        Me.Text13.SelectionStart = 1

        h = 0
        tempstr0 = ""

        If Me.CheckBox1.Checked = True Then
            Me.CheckBox1.BackColor = Color.Red
            'Public PortMitutoyo, PortWeigh As String
            'Pads.SerialPort.PortName = PortWeigh '"COM3"

            Me.Check1.BackColor = Color.Green
            Me.Check1.Checked = False

            If SerialPort.IsOpen Then SerialPort.Close()

            If State = "DE" Then
                Me.CheckBox1.Text = "Mode drücken !"
                Me.Label5.Text = "Waage ID"
                Me.Check1.Text = "Start Messung"
            Else
                Me.CheckBox1.Text = "Press Mode !"
                Me.Label5.Text = "Weight ID"
                Me.Check1.Text = "Start Measurement"
            End If

            Me.Text13.BackColor = Color.Aqua

            If SerialPort.IsOpen = False Then
                Pads.SerialPort.PortName = PortWeigh
                'Pads.SerialPort.PortName = "COM3" ' = .PortName
                'Pads.SerialPort.BaudRate = 9600
                'Pads.SerialPort.DataBits = 8
                'Pads.SerialPort.Parity = Parity.None
                'Pads.SerialPort.StopBits = StopBits.One
                'Pads.SerialPort.Handshake = Handshake.None
            End If

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
                    Me.CheckBox1.Checked = True

                    If State = "DE" Then
                        Me.Text = "Lese auf : " & _
                                                SerialPort.PortName
                    Else
                        Me.Text = "Using port : " & _
                                                SerialPort.PortName
                    End If

                    .RtsEnable = True
                    .DtrEnable = True
                    .ReceivedBytesThreshold = 1

                ElseIf .IsOpen = False Then

                    Me.CheckBox1.Checked = False
                    Me.CheckBox1.BackColor = Color.Green

                    If State = "DE" Then
                        Me.CheckBox1.Text = "Start Wiegen"
                    Else
                        Me.CheckBox1.Text = "Start Weight"
                    End If

                End If

            End With

            If State = "DE" Then
                Me.CheckBox1.Text = "Mode drücken !"
                Me.Label5.Text = "Waage ID"
            Else
                Me.CheckBox1.Text = "Press Mode !"
                Me.Label5.Text = "Weight ID"
            End If

        Else

            If SerialPort.IsOpen Then
                SerialPort.Close()
                Me.CheckBox1.Checked = False

            End If

            tempstr0 = ""

            Me.CheckBox1.BackColor = Color.Green

            If State = "DE" Then
                Me.CheckBox1.Text = "Start Wiegen"
            Else
                Me.CheckBox1.Text = "Start Weight"
            End If

        End If

        If Me.Text13.Text <> "" Then
            Me.Text13.BackColor = Color.White
        Else
            Me.Text13.BackColor = Color.LightSeaGreen
        End If

    End Sub

    Private Sub Check1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Check1.Click
        Dim ex As Exception

        tempstr0 = ""
        ' Checken ob Laufrichtung ausgewaehlt ist:
        ' Check selection of clockside

        If Me.Option1.Checked = False And Me.Option2.Checked = False Then
            If State = "DE" Then
                MsgBox("Bitte Laufrichtung zuerst anwählen !", MsgBoxStyle.Information, "Laufrichtung fehlt")
            Else
                MsgBox("Please Select a Run-Directory !", MsgBoxStyle.Information, "Run-Directory failed")
            End If
            Me.Check1.Checked = False
            Exit Sub
        End If

        ' If checkbox (Start Measurement) checked then

        If Me.Check1.Checked = True Then

            Me.CheckBox1.BackColor = Color.Green
            Me.CheckBox1.Checked = False

            If State = "DE" Then
                Me.CheckBox1.Text = "Start Wiegen"
            Else
                Me.CheckBox1.Text = "Start Weight"
            End If

            If SerialPort.IsOpen Then SerialPort.Close()

        End If

        h = 0
        tempstr0 = ""

        If Me.Check1.Checked = True Then
            Me.Check1.BackColor = Color.Red
            'Public PortMitutoyo, PortWeigh As String
            'Pads.SerialPort.PortName = PortMitutoyo '"COM3"
            Me.Text1.BackColor = Color.Yellow
            Me.CheckBox1.BackColor = Color.Green
            Me.CheckBox1.Checked = False

            If State = "DE" Then
                Me.CheckBox1.Text = "Start Wiegen"
            Else
                Me.CheckBox1.Text = "Start Weight"
            End If

            If SerialPort.IsOpen Then SerialPort.Close()

            If SerialPort.IsOpen = False Then
                Pads.SerialPort.PortName = PortMitutoyo '"COM3"
                'Pads.SerialPort.PortName = "COM3" ' = .PortName
                'Pads.SerialPort.BaudRate = 9600
                'Pads.SerialPort.DataBits = 8
                'Pads.SerialPort.Parity = Parity.None
                'Pads.SerialPort.StopBits = StopBits.One
                'Pads.SerialPort.Handshake = Handshake.None
            End If

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
                    Me.Check1.Checked = True
                    'PortOpenToolStripMenuItem.Checked = True

                    Me.Check1.BackColor = Color.Red

                    If State = "DE" Then
                        Me.Check1.Text = "Lese werte !"
                        Me.Label4.Text = "Messinst. ID"
                        Me.Text = "Lese auf : " & _
                        SerialPort.PortName
                    Else
                        Me.Check1.Text = "Read Values !"
                        Me.Label4.Text = "Meas.Inst. ID"
                        Me.Text = "Using port : " & _
                                                SerialPort.PortName
                    End If

                    'Me.Text = "Using port: " & _
                    'SerialPort.PortName
                    .RtsEnable = True
                    .DtrEnable = True
                    .ReceivedBytesThreshold = 1

                ElseIf .IsOpen = False Then

                    Me.Check1.Checked = False
                    Me.Check1.BackColor = Color.Green

                    If State = "DE" Then
                        Me.Check1.Text = "Start Messung"
                    Else
                        Me.Check1.Text = "Start Measurement"
                    End If
                    'PortOpenToolStripMenuItem.Checked = False
                    'If State = "DE" Then
                    'Me.Text = "keine Daten"
                    'Else
                    'Me.Text = "Not running"
                    'End If

                End If
            End With

        Else

            If SerialPort.IsOpen Then SerialPort.Close()

            Me.Check1.BackColor = Color.Green

            tempstr0 = ""

            If State = "DE" Then
                Me.Check1.Text = "Start Messung"
            Else
                Me.Check1.Text = "Start Measurement"
            End If
        End If

        '#Else
        '#If State = "DE" Then
        '#MsgBox("Bitte Laufrichtung anwählen !", MsgBoxStyle.Information, "Fehlende Auswahl")
        '#Else
        '#MsgBox("Please Select run left or right !", MsgBoxStyle.Information, "Failed Selection")
        '#End If
        '#End If
        '# zum Programmieren ungeeignet : 

        'If SerialPort.IsOpen Then SerialPort.Close()
        'Me.Check1.BackColor = Color.Green
        'tempstr0 = ""
        'If State = "DE" Then
        ' Me.Check1.Text = "Start Messung"
        'Else
        'Me.Check1.Text = "Start Measurement"
        'End If

        Me.TopMost = True

    End Sub

    Public Delegate Sub DisplayData(ByVal Buffer As String)

    Private Shared Sub Display(ByVal Buffer As String)
        Dim i, j, x, y As Integer, valuestr, tmpchr, tmpStr As String

        Buffer = Buffer.Replace(vbLf, vbCr)
        Buffer = Buffer.Replace(vbCr & vbCr, vbCr)
        Buffer = Buffer.Replace(vbCr, vbCrLf)
        y = 0
        h = h + 1

        tempstr1 = ""
        MessInstID = ""

        tempstr0 = tempstr0 & Buffer

        x = Len(tempstr0)

        If Pads.SerialPort.PortName = PortWeigh Then
            If x > 19 Then 'Mid(tempstr0, 19, 1) = "g" Then
                tempstr0 = Mid(tempstr0, 1, 20) 'Trim(tempstr0)
                valuestr = tempstr0
                tmpStr = ""
                tmpchr = ""
                'tempstr0 = ""

                For i = 1 To x
                    tmpchr = Mid(tempstr0, i, 1)
                    If tmpchr <> "+" And tmpchr <> " " And tmpchr <> "(" And _
                       tmpchr <> ")" And tmpchr <> "g" Then
                        tmpStr = tmpStr & tmpchr
                    End If
                Next

                tempstr1 = tmpStr

                If State = "DE" Then
                    For i = 1 To Len(tempstr1)
                        tmpchr = Mid(tempstr1, i, 1)
                        If tmpchr = "." Then
                            tempstr1 = Mid(tempstr1, 1, i - 1) & "," & Mid(tempstr1, i + 1, Len(tempstr1))
                            Exit For
                        End If
                    Next
                End If
                DefInstance.Text13.Text = tempstr1
            End If
        End If


        If Pads.SerialPort.PortName = PortMitutoyo Then

            If h >= 1 And x >= 12 Then
                If x >= 14 Then
                    tempstr0 = Microsoft.VisualBasic.Right(tempstr0, 14)
                    x = Len(tempstr0)
                End If

                valuestr = tempstr0
                tmpStr = ""
                tmpchr = ""
                tempstr0 = ""
                j = 0

                For i = 1 To x

                    tmpchr = Mid(valuestr, i, 1)
                    If tmpchr = "+" And j = 0 Then
                        MessInstID = Mid(tmpStr, 1, 3)
                        tmpStr = ""
                        j = i
                        Exit For
                    End If

                    If tmpchr <> "+" And j = 0 Then
                        tmpStr = Mid(valuestr, 1, i)
                    End If
                Next

                For i = j To x

                    If j > 0 Then
                        If x = 12 Then
                            tmpStr = Mid(valuestr, j + 1, x - j)
                            tempstr1 = tmpStr
                            Exit For
                        End If

                        If x = 13 Then
                            tmpStr = Mid(valuestr, j + 1, x - j - 1)
                            tempstr1 = tmpStr
                            Exit For
                        End If

                        If x = 14 Then
                            tmpStr = Mid(valuestr, j + 1, x - j - 2)
                            tempstr1 = tmpStr
                            Exit For
                        End If
                    End If
                Next
            End If

            tmpStr = ""
            tmpchr = ""

            If tempstr1 <> "" And MessInstID <> "" Then
                If State = "DE" Then
                    For i = 1 To Len(tempstr1)
                        tmpchr = Mid(tempstr1, i, 1)
                        If tmpchr = "." Then
                            tempstr1 = Mid(tempstr1, 1, i - 1) & "," & Mid(tempstr1, i + 1, Len(tempstr1))
                            Exit For
                        End If
                    Next
                End If

                If DefInstance.Text1.Visible = True And DefInstance.Text1.Text = "" Then
                    DefInstance.Text1.Text = tempstr1
                ElseIf DefInstance.Text2.Visible = True And DefInstance.Text2.Text = "" Then
                    DefInstance.Text2.Text = tempstr1
                ElseIf DefInstance.Text3.Visible = True And DefInstance.Text3.Text = "" Then
                    DefInstance.Text3.Text = tempstr1
                ElseIf DefInstance.Text4.Visible = True And DefInstance.Text4.Text = "" Then
                    DefInstance.Text4.Text = tempstr1
                    If DefInstance.Text5.Visible = False And DefInstance.Text4.Text <> "" Then
                        MsgBox(msgMeasurend, MsgBoxStyle.Information)
                        'DefInstance.Check1.Checked = False
                    End If

                ElseIf DefInstance.Text5.Visible = True And DefInstance.Text5.Text = "" Then
                    DefInstance.Text5.Text = tempstr1
                    If DefInstance.Text6.Visible = False And DefInstance.Text5.Text <> "" Then
                        MsgBox(msgMeasurend, MsgBoxStyle.Information)
                        'DefInstance.Check1.Checked = False
                    End If

                ElseIf DefInstance.Text6.Visible = True And DefInstance.Text6.Text = "" Then
                    DefInstance.Text6.Text = tempstr1
                    If DefInstance.Text7.Visible = False And DefInstance.Text6.Text <> "" Then
                        MsgBox(msgMeasurend, MsgBoxStyle.Information)
                        'DefInstance.Check1.Checked = False
                    End If

                ElseIf DefInstance.Text7.Visible = True And DefInstance.Text7.Text = "" Then
                    DefInstance.Text7.Text = tempstr1
                    If DefInstance.Text8.Visible = False And DefInstance.Text7.Text <> "" Then
                        MsgBox(msgMeasurend, MsgBoxStyle.Information)
                        'DefInstance.Check1.Checked = False
                    End If

                ElseIf DefInstance.Text8.Visible = True And DefInstance.Text8.Text = "" Then
                    DefInstance.Text8.Text = tempstr1
                    If DefInstance.Text9.Visible = False And DefInstance.Text8.Text <> "" Then
                        MsgBox(msgMeasurend, MsgBoxStyle.Information)
                        'DefInstance.Check1.Checked = False
                    End If
                ElseIf DefInstance.Text9.Visible = True And DefInstance.Text9.Text = "" Then
                    DefInstance.Text9.Text = tempstr1
                    If DefInstance.Text10.Visible = False And DefInstance.Text9.Text <> "" Then
                        MsgBox(msgMeasurend, MsgBoxStyle.Information)
                        'DefInstance.Check1.Checked = False
                    End If
                ElseIf DefInstance.Text10.Visible = True And DefInstance.Text10.Text = "" Then
                    DefInstance.Text10.Text = tempstr1
                    If DefInstance.Text11.Visible = False And DefInstance.Text10.Text <> "" Then
                        MsgBox(msgMeasurend, MsgBoxStyle.Information)
                        'DefInstance.Check1.Checked = False
                    End If
                ElseIf DefInstance.Text11.Visible = True And DefInstance.Text11.Text = "" Then
                    DefInstance.Text11.Text = tempstr1
                    If DefInstance.Text12.Visible = False And DefInstance.Text11.Text <> "" Then
                        MsgBox(msgMeasurend, MsgBoxStyle.Information)
                        'DefInstance.Check1.Checked = False
                    End If
                ElseIf DefInstance.Text12.Visible = True And DefInstance.Text12.Text = "" Then
                    DefInstance.Text12.Text = tempstr1
                    MsgBox(msgMeasurend, MsgBoxStyle.Information)
                    'DefInstance.Check1.Checked = False
                End If

                tempstr1 = ""

                '* If DefInstance.TextBox3.Text = "" Then
                '* tempstr1 = Microsoft.VisualBasic.Left(MessInstID, 2)
                '* y = CInt(Microsoft.VisualBasic.Right(tempstr1, 1))
                '* If y = 1 Then
                'Public MitutoyoInstr1ID, MitutoyoInstr2ID, weightID As Integer
                '*     DefInstance.TextBox3.Text = MitutoyoInstr1ID
                '* ElseIf y = 2 Then
                '*     DefInstance.TextBox3.Text = MitutoyoInstr2ID
                '*  End If
                'DefInstance.TextBox3.Text = MessInstID
                '*End If
                'DefInstance.TextBox3.Text = DefInstance.TextBox4.Text
                '* tempstr1 = ""
                '* MessInstID = ""
            End If
            DefInstance.TextBox3.Text = DefInstance.TextBox4.Text
            MessInstID = ""
        End If

    End Sub

    '#

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        If Me.Option1.BackColor = Color.DarkRed Then
            If State = "DE" Then
                MsgBox("Sie müssen die Laufrichtung auswählen/bestätigen !")
                Exit Sub
            Else
                MsgBox("Select a side of pad !")
                Exit Sub
            End If
        End If

        '#Me.TopMost = False

        canchelPad = 1

        Me.TextBox4.Text = Me.ComboBox1.Text

        If Me.TextBox4.Text = "" Then
            If State = "DE" Then
                MsgBox("Sie müssen die Messinstr. ID eintragen !")
                Exit Sub
            Else
                MsgBox("Insert a Measurement Instr. ID !")
                Exit Sub
            End If
        End If

        Me.TextBox2.Text = Me.ComboBox2.Text

        If Me.TextBox2.Text = "" Then
            If State = "DE" Then
                MsgBox("Sie müssen die Waagen ID eintragen !")
                Exit Sub
            Else
                MsgBox("Insert a Weigh Instr. ID !")
                Exit Sub
            End If
        End If

        If DBOrExcel = 2 Then
            ValuesWritetoexcel()
        ElseIf DBOrExcel = 1 Then
            ValuesToDB()
        End If

        Me.Close()
        Form2.Show()

    End Sub

    '#################################################################################
    '# Set here the color, selection, of actual measurepoint:
    '# Hier wird die Farbe gesetzt, die den aktuellen Messpunkt kennzeichnet:
    '#################################################################################

    Private Sub Text1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text1.TextChanged
        If Me.Text1.Text <> "" Then
            Me.Text1.BackColor = Color.White
            If Me.Text2.Visible = True Then
                'Me.Text2.BackColor = Color.LightSeaGreen
                Me.Text2.BackColor = Color.Yellow
            ElseIf Me.Text3.Visible = True And Me.Text3.Text = "" _
                   And Me.Text2.Visible = False Then
                Me.Text3.BackColor = Color.Yellow
                '      Me.Text3.BackColor = Color.LightSeaGreen
            End If
        Else
            Me.Text1.BackColor = Color.Yellow
            '  Me.Text1.BackColor = Color.LightSeaGreen
        End If
    End Sub

    Private Sub Text2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text2.TextChanged
        If Me.Text2.Text <> "" Then
            Me.Text2.BackColor = Color.White
            If Me.Text3.Visible = True Then
                Me.Text3.BackColor = Color.Yellow
                'Me.Text3.BackColor = Color.LightSeaGreen
            ElseIf Me.Text4.Visible = True And Me.Text4.Text = "" _
                   And Me.Text3.Visible = False Then
                Me.Text4.BackColor = Color.Yellow
                '      Me.Text4.BackColor = Color.LightSeaGreen
            End If
        Else
            Me.Text2.BackColor = Color.Yellow
            '  Me.Text2.BackColor = Color.LightSeaGreen
        End If
    End Sub


    Private Sub Text3_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text3.TextChanged
        If Me.Text3.Text <> "" Then
            Me.Text3.BackColor = Color.White
            If Me.Text4.Visible = True Then
                'Me.Text4.BackColor = Color.LightSeaGreen
                Me.Text4.BackColor = Color.Yellow
            ElseIf Me.Text5.Visible = True And Me.Text5.Text = "" _
                   And Me.Text4.Visible = False Then
                Me.Text5.BackColor = Color.Yellow
                '     Me.Text5.BackColor = Color.LightSeaGreen
            End If
        Else
            Me.Text3.BackColor = Color.Yellow
            ' Me.Text3.BackColor = Color.LightSeaGreen
        End If
    End Sub

    Private Sub Text4_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text4.TextChanged
        If Me.Text4.Text <> "" Then
            Me.Text4.BackColor = Color.White
            If Me.Text5.Visible = True Then
                'Me.Text5.BackColor = Color.LightSeaGreen
                Me.Text5.BackColor = Color.Yellow
            ElseIf Me.Text6.Visible = True And Me.Text6.Text = "" _
                   And Me.Text5.Visible = False Then
                Me.Text6.BackColor = Color.Yellow
                '    Me.Text6.BackColor = Color.LightSeaGreen
            End If
        Else
            Me.Text4.BackColor = Color.Yellow
            ' Me.Text4.BackColor = Color.LightSeaGreen
        End If
    End Sub

    Private Sub Text5_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text5.TextChanged
        If Me.Text5.Text <> "" Then
            Me.Text5.BackColor = Color.White
            If Me.Text6.Visible = True Then
                'Me.Text6.BackColor = Color.LightSeaGreen
                Me.Text6.BackColor = Color.Yellow
            ElseIf Me.Text7.Visible = True And Me.Text7.Text = "" _
                   And Me.Text6.Visible = False Then
                Me.Text7.BackColor = Color.Yellow
                '     Me.Text7.BackColor = Color.LightSeaGreen
            End If
        Else
            Me.Text5.BackColor = Color.Yellow
            ' Me.Text5.BackColor = Color.LightSeaGreen
        End If
    End Sub

    Private Sub Text6_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text6.TextChanged
        If Me.Text6.Text <> "" Then
            Me.Text6.BackColor = Color.White
            If Me.Text7.Visible = True Then
                'Me.Text7.BackColor = Color.LightSeaGreen
                Me.Text7.BackColor = Color.Yellow
            ElseIf Me.Text8.Visible = True And Me.Text8.Text = "" _
                   And Me.Text7.Visible = False Then
                Me.Text8.BackColor = Color.Yellow
                '    Me.Text8.BackColor = Color.LightSeaGreen
            End If
        Else
            Me.Text6.BackColor = Color.Yellow
            'Me.Text6.BackColor = Color.LightSeaGreen
        End If
    End Sub

    Private Sub Text7_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text7.TextChanged
        If Me.Text7.Text <> "" Then
            Me.Text7.BackColor = Color.White
            If Me.Text8.Visible = True Then
                Me.Text8.BackColor = Color.Yellow
                'Me.Text8.BackColor = Color.LightSeaGreen
            ElseIf Me.Text9.Visible = True And Me.Text9.Text = "" _
                   And Me.Text8.Visible = False Then
                Me.Text9.BackColor = Color.Yellow
                '   Me.Text9.BackColor = Color.LightSeaGreen
            End If
        Else
            Me.Text7.BackColor = Color.Yellow
            'Me.Text7.BackColor = Color.LightSeaGreen
        End If
    End Sub

    Private Sub Text8_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text8.TextChanged
        If Me.Text8.Text <> "" Then
            Me.Text8.BackColor = Color.White
            If Me.Text9.Visible = True Then
                Me.Text9.BackColor = Color.Yellow
                'Me.Text9.BackColor = Color.LightSeaGreen
            ElseIf Me.Text10.Visible = True And Me.Text10.Text = "" _
                   And Me.Text9.Visible = False Then
                Me.Text10.BackColor = Color.Yellow
                'Me.Text10.BackColor = Color.LightSeaGreen
            End If
        Else
            Me.Text8.BackColor = Color.Yellow
            'Me.Text8.BackColor = Color.LightSeaGreen
        End If
    End Sub

    Private Sub Text9_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text9.TextChanged
        If Me.Text9.Text <> "" Then
            Me.Text9.BackColor = Color.White
            If Me.Text10.Visible = True Then
                'Me.Text10.BackColor = Color.LightSeaGreen
                Me.Text10.BackColor = Color.Yellow
            ElseIf Me.Text11.Visible = True And Me.Text11.Text = "" _
                   And Me.Text10.Visible = False Then
                Me.Text11.BackColor = Color.Yellow
                '  Me.Text11.BackColor = Color.LightSeaGreen
            End If
        Else
            Me.Text9.BackColor = Color.Yellow
            'Me.Text9.BackColor = Color.LightSeaGreen
        End If
    End Sub

    Private Sub Text10_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text10.TextChanged
        If Me.Text10.Text <> "" Then
            Me.Text10.BackColor = Color.White
            If Me.Text11.Visible = True Then
                Me.Text11.BackColor = Color.Yellow
                'Me.Text11.BackColor = Color.LightSeaGreen
            ElseIf Me.Text12.Visible = True And Me.Text12.Text = "" _
                   And Me.Text11.Visible = False Then
                Me.Text12.BackColor = Color.Yellow
                '    Me.Text12.BackColor = Color.LightSeaGreen
            End If
        Else
            Me.Text10.BackColor = Color.Yellow
            'Me.Text10.BackColor = Color.LightSeaGreen
        End If
    End Sub

    Private Sub Text11_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text11.TextChanged
        If Me.Text11.Text <> "" Then
            Me.Text11.BackColor = Color.White
            If Me.Text12.Visible = True Then
                Me.Text12.BackColor = Color.Yellow
            ElseIf Me.Text13.Visible = True And Me.Text13.Text = "" _
                   And Me.Text12.Visible = False Then
                Me.Text13.BackColor = Color.Yellow
            End If
        Else
            Me.Text11.BackColor = Color.Yellow
        End If
    End Sub

    Private Sub Text12_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text12.TextChanged
        If Me.Text12.Text <> "" Then
            Me.Text12.BackColor = Color.White
        Else
            Me.Text12.BackColor = Color.LightSeaGreen
        End If
    End Sub

    Private Sub Text13_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text13.TextChanged
        If Me.Text13.Text <> "" Then
            Me.Text13.BackColor = Color.White
            Me.CheckBox1.Checked = False
        Else
            Me.Text13.BackColor = Color.LightSeaGreen
        End If
    End Sub

    Private Sub Option1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Option1.CheckedChanged

        lngNumCols = CInt(Form2.Label2.Text)

        If PadsDir = "updown" Then

            If Me.Option1.Checked = True And lngNumCols = 4 Then

                Me.Option2.BackColor = Color.Gray
                Me.Option1.BackColor = Color.Green

                Me.Text1.Location = New Point(8, 93)
                Me.Label7.Location = New Point(39, 77)
                Me.Label7.Visible = True
                Me.Label7.Text = "1"

                Me.Text2.Location = New Point(256, 24)
                Me.Label8.Location = New Point(49, 172)
                Me.Label8.Visible = True
                Me.Label8.Text = "2"

            ElseIf Me.Option1.Checked = True And lngNumCols = 6 Then

                Me.Option2.BackColor = Color.Gray
                Me.Option1.BackColor = Color.Green

                Me.Text1.Location = New Point(8, 93)
                Me.Label7.Location = New Point(39, 77)
                Me.Label7.Visible = True
                Me.Label7.Text = "1"
                Me.Text2.Location = New Point(38, 189)

                Me.Text3.Location = New Point(256, 24)
                Me.Label8.Location = New Point(49, 172)
                Me.Label8.Visible = True
                Me.Label8.Text = "2"
                Me.Text4.Location = New Point(256, 148)

                Me.Text5.Location = New Point(430, 93)
                Me.Label9.Location = New Point(253, 8)
                Me.Label9.Visible = True
                Me.Label9.Text = "3"
                Me.Text6.Location = New Point(404, 189)

            ElseIf Me.Option1.Checked = True And lngNumCols = 9 Then

                Me.Option2.BackColor = Color.Gray
                Me.Option1.BackColor = Color.Green

                Me.Text1.Location = New Point(8, 93)
                Me.Label7.Location = New Point(39, 77)
                Me.Label7.Visible = True
                Me.Label7.Text = "1"


                Me.Text2.Location = New Point(8, 146)
                Me.Label8.Location = New Point(19, 130)
                Me.Label8.Visible = True
                Me.Label8.Text = "2"

                Me.Text3.Location = New Point(38, 189)
                Me.Label9.Location = New Point(49, 172)
                Me.Label9.Visible = True
                Me.Label9.Text = "3"

                Me.Text4.Location = New Point(256, 24)
                Me.Text5.Location = New Point(256, 77)
                Me.Text6.Location = New Point(256, 148)

                Me.Text7.Location = New Point(430, 93)
                Me.Text8.Location = New Point(430, 146)
                Me.Text9.Location = New Point(404, 189)

            End If
        ElseIf PadsDir = "round" Then

            If Me.Option1.Checked = True And lngNumCols = 4 Then

                Me.Option2.BackColor = Color.Gray
                Me.Option1.BackColor = Color.Green

                Me.Text1.Location = New Point(8, 93)
                Me.Label7.Location = New Point(39, 77)
                Me.Label7.Visible = True
                Me.Label7.Text = "1"

                Me.Text2.Location = New Point(430, 93)
                Me.Label9.Location = New Point(427, 77)
                Me.Label9.Visible = True
                Me.Label9.Text = "2"

                Me.Text3.Location = New Point(38, 189)
                Me.Text4.Location = New Point(404, 189)

            ElseIf Me.Option1.Checked = True And lngNumCols = 6 Then

                Me.Option2.BackColor = Color.Gray
                Me.Option1.BackColor = Color.Green

                Me.Text1.Location = New Point(8, 93)
                Me.Label7.Location = New Point(39, 77)
                Me.Label7.Visible = True
                Me.Label7.Text = "1"

                Me.Text2.Location = New Point(256, 24)
                Me.Label8.Location = New Point(253, 8)
                Me.Label8.Visible = True
                Me.Label8.Text = "2"

                Me.Text3.Location = New Point(430, 93)
                Me.Label9.Location = New Point(427, 77)
                Me.Label9.Visible = True
                Me.Label9.Text = "3"

                Me.Text4.Location = New Point(38, 189)
                Me.Text5.Location = New Point(256, 148)
                Me.Text6.Location = New Point(404, 189)

            ElseIf Me.Option1.Checked = True And lngNumCols = 9 Then

                Me.Option2.BackColor = Color.Gray
                Me.Option1.BackColor = Color.Green

                Me.Text1.Location = New Point(8, 93)
                Me.Label7.Location = New Point(39, 77)
                Me.Label7.Visible = True
                Me.Label7.Text = "1"

                Me.Text2.Location = New Point(256, 24)
                Me.Label8.Location = New Point(253, 8)
                Me.Label8.Visible = True
                Me.Label8.Text = "2"

                Me.Text3.Location = New Point(430, 93)
                Me.Label9.Location = New Point(427, 77)
                Me.Label9.Visible = True
                Me.Label9.Text = "3"

                Me.Text4.Location = New Point(8, 146)
                Me.Text5.Location = New Point(256, 77)
                Me.Text6.Location = New Point(430, 146)

                Me.Text7.Location = New Point(38, 189)
                Me.Text8.Location = New Point(256, 148)
                Me.Text9.Location = New Point(404, 189)

            End If
        End If

    End Sub

    Private Sub Option2_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Option2.CheckedChanged

        lngNumCols = CInt(Form2.Label2.Text)

        If PadsDir = "updown" Then

            If Me.Option2.Checked = True And lngNumCols = 6 Then

                Me.Option1.BackColor = Color.Gray
                Me.Option2.BackColor = Color.Green

                Me.Text5.Location = New Point(8, 93)
                Me.Text6.Location = New Point(38, 189)

                Me.Text3.Location = New Point(256, 24)
                Me.Label9.Location = New Point(267, 8)
                Me.Label9.Visible = True
                Me.Label9.Text = "3"

                Me.Text4.Location = New Point(256, 148)

                Me.Text1.Location = New Point(430, 93)
                Me.Label7.Location = New Point(427, 77)
                Me.Label7.Visible = True
                Me.Label7.Text = "1"

                Me.Text2.Location = New Point(403, 189)
                Me.Label8.Location = New Point(403, 173)
                Me.Label8.Visible = True
                Me.Label8.Text = "2"


            ElseIf Me.Option2.Checked = True And lngNumCols = 9 Then

                Me.Option1.BackColor = Color.Gray
                Me.Option2.BackColor = Color.Green

                Me.Text7.Location = New Point(8, 93)
                Me.Text8.Location = New Point(8, 146)
                Me.Text9.Location = New Point(38, 189)

                Me.Text4.Location = New Point(256, 24)
                Me.Text5.Location = New Point(256, 77)
                Me.Text6.Location = New Point(256, 148)

                Me.Text1.Location = New Point(430, 93)
                Me.Label7.Location = New Point(427, 77)
                Me.Label7.Visible = True
                Me.Label7.Text = "1"

                Me.Text2.Location = New Point(430, 146)
                Me.Label8.Location = New Point(427, 130)
                Me.Label8.Visible = True
                Me.Label8.Text = "2"

                Me.Text3.Location = New Point(404, 189)
                Me.Label9.Location = New Point(403, 173)
                Me.Label9.Visible = True
                Me.Label9.Text = "3"
            End If

        ElseIf PadsDir = "round" Then

            If Me.Option2.Checked = True And lngNumCols = 4 Then

                Me.Option1.BackColor = Color.Gray
                Me.Option2.BackColor = Color.Green

                Me.Text2.Location = New Point(8, 93)

                Me.Label7.Location = New Point(39, 77)
                Me.Label7.Visible = True
                Me.Label7.Text = "2"

                Me.Text1.Location = New Point(430, 93)

                Me.Label9.Location = New Point(427, 77)
                Me.Label9.Visible = True
                Me.Label9.Text = "1"

                Me.Text4.Location = New Point(38, 189)
                Me.Text3.Location = New Point(404, 189)

            ElseIf Me.Option2.Checked = True And lngNumCols = 6 Then

                Me.Option1.BackColor = Color.Gray
                Me.Option2.BackColor = Color.Green

                Me.Text3.Location = New Point(8, 93)
                Me.Text6.Location = New Point(38, 189)

                Me.Text2.Location = New Point(256, 24)
                Me.Text5.Location = New Point(256, 148)

                Me.Text1.Location = New Point(430, 93)
                Me.Label7.Location = New Point(427, 77)
                Me.Label7.Visible = True
                Me.Label7.Text = "1"

                Me.Label8.Location = New Point(253, 8)
                Me.Label8.Visible = True
                Me.Label8.Text = "2"

                Me.Label9.Location = New Point(39, 77)
                Me.Label9.Visible = True
                Me.Label9.Text = "3"

                Me.Text4.Location = New Point(404, 189)

            ElseIf Me.Option2.Checked = True And lngNumCols = 9 Then

                Me.Option1.BackColor = Color.Gray
                Me.Option2.BackColor = Color.Green

                Me.Text3.Location = New Point(8, 93)
                Me.Text6.Location = New Point(8, 146)
                Me.Text9.Location = New Point(38, 189)

                Me.Text2.Location = New Point(256, 24)
                Me.Text5.Location = New Point(256, 77)
                Me.Text8.Location = New Point(256, 148)

                Me.Text1.Location = New Point(430, 93)
                Me.Text4.Location = New Point(430, 146)
                Me.Text7.Location = New Point(404, 189)

                Me.Label7.Location = New Point(427, 77)
                Me.Label7.Visible = True
                Me.Label7.Text = "1"

                Me.Label8.Location = New Point(253, 8)
                Me.Label8.Visible = True
                Me.Label8.Text = "2"

                Me.Label9.Location = New Point(39, 77)
                Me.Label9.Visible = True
                Me.Label9.Text = "3"

            End If
        End If

    End Sub

    Private Sub Option1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Option1.Click
        If Me.Option1.Checked = True Then
            Me.Option1.BackColor = Color.Green
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        ValuesWritetoexcel()
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If Me.CheckBox2.Checked = True Then
            Me.Width = 768
            Me.Height = 351
        Else
            Me.Width = 640
            Me.Height = 351
        End If
    End Sub

    Private Sub ComboBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.TextChanged
        If Me.ComboBox1.Text <> "" Then
            Me.ComboBox1.BackColor = Color.White
        Else
            Me.ComboBox1.BackColor = Color.LightSeaGreen
        End If

    End Sub

    Private Sub ComboBox2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.TextChanged
        If Me.ComboBox2.Text <> "" Then
            Me.ComboBox2.BackColor = Color.White
        Else
            Me.ComboBox2.BackColor = Color.LightSeaGreen
        End If
    End Sub

End Class

