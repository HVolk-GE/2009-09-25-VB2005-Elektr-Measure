Option Strict Off
Option Explicit On
Imports System.IO.Ports

Friend Class Disc
    Inherits System.Windows.Forms.Form
    Public Shared WithEvents SerialPort As SerialPort
    Private Shared m_FormDefInstance As Disc
    Private Shared m_InitializingDefInstance As Boolean
    Public Measurepoint As Integer
    Private dataSet As DataSet

    Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
        Me.Close()
        Form2.Show()
    End Sub

    Public Sub Disc_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        On Error Resume Next

        If SerialPort.IsOpen Then SerialPort.Close()

    End Sub

    Private Shared Sub SerialPort_DataReceived(ByVal sender As Object, ByVal e As _
       System.IO.Ports.SerialDataReceivedEventArgs) Handles SerialPort.DataReceived
        Dim Buffer As String = SerialPort.ReadExisting()
        DefInstance.Text1.BeginInvoke(New _
            DisplayData(AddressOf Display), _
                    New Object() {Buffer})
    End Sub

    Public Shared Property DefInstance() As Disc
        Get
            If m_FormDefInstance Is Nothing OrElse _
                        m_FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_FormDefInstance = New Disc
                m_InitializingDefInstance = False
            End If
            DefInstance = m_FormDefInstance
        End Get
        Set(ByVal Value As Disc)
            m_FormDefInstance = Value
        End Set
    End Property

    Private Sub Disc_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        Measurepoint = 1
        SerialPort = New SerialPort
        DefInstance = Me

        Me.Check1.BackColor = Color.Green
        Me.CheckBox1.BackColor = Color.Green
        Me.ComboBox1.Text = ""
        Me.ComboBox2.Text = ""

        Me.Width = 773
        Me.Height = 612

        If State = "DE" Then
            Me.Label3.Text = "Kontrolle :"
            Me.Check1.Text = "Start Messung"
            Me.Label1.Text = "Benutzername"
            Me.Label2.Text = "Gewicht (g)"
            Me.Command1.Text = "Abbrechen"
            Me.Button1.Text = "OK"
            Me.Label4.Text = "Messinst. ID"
            Me.Label5.Text = "Waage ID"
            Me.Label20.Text = "Disc stat."
            Me.CheckBox2.Text = "Edit Disc ID"
            Me.CheckBox1.Text = "Start Wiegen"
            msgMeasurend = "Anzahl der Messpunkte erreicht, Messung beendet !"
        Else
            Me.Label3.Text = "Control :"
            Me.Label1.Text = "Username"
            Me.Label2.Text = "Weight (g)"
            Me.Command1.Text = "Cancel"
            Me.Button1.Text = "OK"
            Me.Check1.Text = "Start Measurement"
            Me.Label4.Text = "Meas.Inst. ID"
            Me.Label5.Text = "Weigh ID"
            Me.Label20.Text = "Disc cond."
            Me.CheckBox2.Text = "Edit Disc ID"
            Me.CheckBox1.Text = "Start Weigh"
            msgMeasurend = "Measruement finished !"
        End If

        Me.Text19.Text = strUsername
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
        Me.ComboBox1.BackColor = Color.LightSeaGreen
        Me.ComboBox2.BackColor = Color.LightSeaGreen

        If Me.TextBox4.Text = "" Then
            Me.TextBox4.BackColor = Color.LightSeaGreen
            Me.TextBox4.Enabled = True
            Me.CheckBox2.Enabled = False
        Else
            Me.TextBox4.BackColor = Color.White
            Me.TextBox4.Enabled = False
            Me.CheckBox2.Enabled = True
        End If

        Me.TextBox5.BackColor = Color.LightSeaGreen
        Me.TextBox6.BackColor = Color.LightSeaGreen

        If Me.ComboBox3.Text = "" Then Me.ComboBox3.BackColor = Color.LightSeaGreen

        Me.TextBox2.BackColor = Color.LightSeaGreen
        Me.TextBox2.Text = ""
        Me.TextBox2.Enabled = True
        Me.CheckBox2.Checked = False

        Me.ComboBox1.Items.Add(MitutoyoInstr1ID)
        Me.ComboBox1.Items.Add(MitutoyoInstr2ID)
        Me.ComboBox1.Items.Add(MitutoyoInstr3ID)
        Me.ComboBox1.Items.Add(MitutoyoInstr4ID)

        Me.ComboBox2.Items.Add(weightID)
        Me.ComboBox2.Items.Add(weight1ID)

        Me.TopMost = True

        Exit Sub

err_Handler:
        'Wenn ein Fehler aufgetreten ist, alle Controls
        'ausser Beenden-Button deaktivieren:
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
            If SerialPort.IsOpen Then SerialPort.Close()
            tempstr0 = ""
            Me.Check1.BackColor = Color.Green
            Me.Check1.Checked = False

            If State = "DE" Then
                Me.Check1.Text = "Start Messung"
            Else
                Me.Check1.Text = "Start Measurement"
            End If
        End If

        Me.Text13.SelectionStart = 1

        h = 0
        tempstr0 = ""

        If Me.CheckBox1.Checked = True Then
            Me.CheckBox1.BackColor = Color.Red
            Me.Text13.BackColor = Color.Yellow

            If SerialPort.IsOpen = False Then
                Disc.SerialPort.PortName = PortWeigh '"COM3"
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
                Else
                    Me.CheckBox1.Checked = False
                    If State = "DE" Then
                        Me.Text = "Lese keine Daten"
                    Else
                        Me.Text = "Not running"
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
            If SerialPort.IsOpen Then SerialPort.Close()
            Me.CheckBox1.BackColor = Color.Green
            tempstr0 = ""
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

        MessInstID = ""
        h = 0
        tempstr0 = ""

        If Me.CheckBox1.Checked = True Then
            Me.CheckBox1.BackColor = Color.Green
            Me.CheckBox1.Checked = False
            If State = "DE" Then
                Me.CheckBox1.Text = "Start Wiegen"
            Else
                Me.CheckBox1.Text = "Start Weight"
            End If
            If SerialPort.IsOpen Then SerialPort.Close()
        End If

        If Me.Check1.Checked = True Then
            Me.Check1.BackColor = Color.Red
            Me.Text1.BackColor = Color.Yellow

            If SerialPort.IsOpen = False Then
                Disc.SerialPort.PortName = PortMitutoyo 'Here USB-Box = "COM3"
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
                Else
                    Me.Check1.Checked = False
                    If State = "DE" Then
                        Me.Text = "Lese keine Daten"
                    Else
                        Me.Text = "Not running"
                    End If

                End If
            End With
            If State = "DE" Then
                Me.Check1.Text = "lese Werte !"
                Me.Label4.Text = "Messinst. ID"
            Else
                Me.Check1.Text = "Read Values !"
                Me.Label4.Text = "Meas.Inst. ID"
            End If

        Else
            If SerialPort.IsOpen Then SerialPort.Close()
            tempstr0 = ""
            Me.Check1.BackColor = Color.Green
            If State = "DE" Then
                Me.Check1.Text = "Start Messung"
            Else
                Me.Check1.Text = "Start Measurement"
            End If
        End If

    End Sub

    Public Delegate Sub DisplayData(ByVal Buffer As String)

    Private Shared Sub Display(ByVal Buffer As String)
        Dim i, j, x, y, valu As Integer, valuestr, tmpchr, tmpStr As String

        Buffer = Buffer.Replace(vbLf, vbCr)
        Buffer = Buffer.Replace(vbCr & vbCr, vbCr)
        Buffer = Buffer.Replace(vbCr, vbCrLf)
        y = 0
        h = h + 1
        Dim iDbl, jDbl, xDbl As Double
        iDbl = 0
        jDbl = 0
        xDbl = 0

        tempstr1 = ""

        tempstr0 = tempstr0 & Buffer

        x = Len(tempstr0)

        If Disc.SerialPort.PortName = PortWeigh Then
            If x > 19 Then
                tempstr0 = Mid(tempstr0, 1, 20)
                'String from Weigh: "    +    22.(3)    g"
                valuestr = tempstr0
                tmpStr = ""
                tmpchr = ""

                For i = 1 To x
                    tmpchr = Mid(tempstr0, i, 1)

                    If tmpchr <> Chr0 And tmpchr <> Chr1 And tmpchr <> Chr2 And tmpchr <> Chr3 And _
                       tmpchr <> Chr4 Then
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

        If Disc.SerialPort.PortName = PortMitutoyo Then

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

                    valu = CInt(tempstr1)
                    iDbl = CDbl(valu) - 5.0
                    jDbl = CDbl(valu) + 5.0

                ElseIf DefInstance.Text2.Visible = True And DefInstance.Text2.Text = "" Then
                    DefInstance.Text2.Text = tempstr1
                ElseIf DefInstance.Text3.Visible = True And DefInstance.Text3.Text = "" Then
                    DefInstance.Text3.Text = tempstr1
                ElseIf DefInstance.Text4.Visible = True And DefInstance.Text4.Text = "" Then
                    DefInstance.Text4.Text = tempstr1
                ElseIf DefInstance.Text5.Visible = True And DefInstance.Text5.Text = "" Then
                    DefInstance.Text5.Text = tempstr1
                ElseIf DefInstance.Text6.Visible = True And DefInstance.Text6.Text = "" Then
                    DefInstance.Text6.Text = tempstr1
                ElseIf DefInstance.Text7.Visible = True And DefInstance.Text7.Text = "" Then
                    DefInstance.Text7.Text = tempstr1
                ElseIf DefInstance.Text8.Visible = True And DefInstance.Text8.Text = "" Then
                    DefInstance.Text8.Text = tempstr1
                ElseIf DefInstance.Text9.Visible = True And DefInstance.Text9.Text = "" Then
                    DefInstance.Text9.Text = tempstr1
                ElseIf DefInstance.Text10.Visible = True And DefInstance.Text10.Text = "" Then
                    DefInstance.Text10.Text = tempstr1
                ElseIf DefInstance.Text11.Visible = True And DefInstance.Text11.Text = "" Then
                    DefInstance.Text11.Text = tempstr1
                ElseIf DefInstance.Text12.Visible = True And DefInstance.Text12.Text = "" Then
                    DefInstance.Text12.Text = tempstr1
                    MsgBox(msgMeasurend, MsgBoxStyle.Information)
                End If

                tempstr1 = ""

                If DefInstance.TextBox3.Text = "" Then
                    tempstr1 = Microsoft.VisualBasic.Left(MessInstID, 2)
                    y = CInt(Microsoft.VisualBasic.Right(tempstr1, 1))
                    If y = 1 Then
                        DefInstance.TextBox3.Text = MitutoyoInstr1ID
                    ElseIf y = 2 Then
                        DefInstance.TextBox3.Text = MitutoyoInstr2ID
                    End If

                End If
                tempstr1 = ""
            End If

        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        Me.TextBox6.Text = Me.ComboBox1.Text

        If Me.TextBox6.Text = "" Then
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

        Me.TextBox3.Text = Me.TextBox6.Text

        If DBOrExcel = 2 Then
            ValuesWritetoexcel()
        ElseIf DBOrExcel = 1 Then
            ValuesToDB()
        End If

        canchelPad = 0

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
                Me.Text2.BackColor = Color.Yellow
            ElseIf Me.Text3.Visible = True And Me.Text3.Text = "" _
                   And Me.Text2.Visible = False Then
                Me.Text3.BackColor = Color.Yellow
            End If
        Else
            Me.Text1.BackColor = Color.Yellow
        End If
    End Sub

    Private Sub Text2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text2.TextChanged
        If Me.Text2.Text <> "" Then
            Me.Text2.BackColor = Color.White
            If Me.Text3.Visible = True Then
                Me.Text3.BackColor = Color.Yellow
            ElseIf Me.Text4.Visible = True And Me.Text4.Text = "" _
                   And Me.Text3.Visible = False Then
                Me.Text4.BackColor = Color.Yellow
            End If
        Else
            Me.Text2.BackColor = Color.Yellow
        End If
    End Sub

    Private Sub Text3_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text3.TextChanged
        If Me.Text3.Text <> "" Then
            Me.Text3.BackColor = Color.White
            If Me.Text4.Visible = True Then
                Me.Text4.BackColor = Color.Yellow
            ElseIf Me.Text5.Visible = True And Me.Text5.Text = "" _
                   And Me.Text4.Visible = False Then
                Me.Text5.BackColor = Color.Yellow
            End If
        Else
            Me.Text3.BackColor = Color.Yellow
        End If
    End Sub

    Private Sub Text4_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text4.TextChanged
        If Me.Text4.Text <> "" Then
            Me.Text4.BackColor = Color.White
            If Me.Text5.Visible = True Then
                Me.Text5.BackColor = Color.Yellow
            ElseIf Me.Text6.Visible = True And Me.Text6.Text = "" _
                   And Me.Text5.Visible = False Then
                Me.Text6.BackColor = Color.Yellow
            End If
        Else
            Me.Text4.BackColor = Color.Yellow
        End If
    End Sub

    Private Sub Text5_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text5.TextChanged
        If Me.Text5.Text <> "" Then
            Me.Text5.BackColor = Color.White
            If Me.Text6.Visible = True Then
                Me.Text6.BackColor = Color.Yellow
            ElseIf Me.Text7.Visible = True And Me.Text7.Text = "" _
                   And Me.Text6.Visible = False Then
                Me.Text7.BackColor = Color.Yellow
            End If
        Else
            Me.Text5.BackColor = Color.Yellow
        End If
    End Sub

    Private Sub Text6_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text6.TextChanged
        If Me.Text6.Text <> "" Then
            Me.Text6.BackColor = Color.White
            If Me.Text7.Visible = True Then
                Me.Text7.BackColor = Color.Yellow
            ElseIf Me.Text8.Visible = True And Me.Text8.Text = "" _
                   And Me.Text7.Visible = False Then
                Me.Text8.BackColor = Color.Yellow
            End If
        Else
            Me.Text6.BackColor = Color.Yellow
        End If
    End Sub

    Private Sub Text7_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text7.TextChanged
        If Me.Text7.Text <> "" Then
            Me.Text7.BackColor = Color.White
            If Me.Text8.Visible = True Then
                Me.Text8.BackColor = Color.Yellow
            ElseIf Me.Text9.Visible = True And Me.Text9.Text = "" _
                   And Me.Text8.Visible = False Then
                Me.Text9.BackColor = Color.Yellow
            End If
        Else
            Me.Text7.BackColor = Color.Yellow
        End If
    End Sub

    Private Sub Text8_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text8.TextChanged
        If Me.Text8.Text <> "" Then
            Me.Text8.BackColor = Color.White
            If Me.Text9.Visible = True Then
                Me.Text9.BackColor = Color.Yellow
            ElseIf Me.Text10.Visible = True And Me.Text10.Text = "" _
                   And Me.Text9.Visible = False Then
                Me.Text10.BackColor = Color.Yellow
            End If
        Else
            Me.Text8.BackColor = Color.Yellow
        End If
    End Sub

    Private Sub Text9_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text9.TextChanged
        If Me.Text9.Text <> "" Then
            Me.Text9.BackColor = Color.White
            If Me.Text10.Visible = True Then
                Me.Text10.BackColor = Color.Yellow
            ElseIf Me.Text11.Visible = True And Me.Text11.Text = "" _
                   And Me.Text10.Visible = False Then
                Me.Text11.BackColor = Color.Yellow
            End If
        Else
            Me.Text9.BackColor = Color.Yellow
        End If
    End Sub

    Private Sub Text10_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text10.TextChanged
        If Me.Text10.Text <> "" Then
            Me.Text10.BackColor = Color.White
            If Me.Text11.Visible = True Then
                Me.Text11.BackColor = Color.Yellow
            ElseIf Me.Text12.Visible = True And Me.Text12.Text = "" _
                   And Me.Text11.Visible = False Then
                Me.Text12.BackColor = Color.Yellow
            End If
        Else
            Me.Text10.BackColor = Color.Yellow
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
        Else
            Me.Text13.BackColor = Color.LightSeaGreen
        End If
    End Sub
    '##############################################################################################
    Sub FillDatas()
        Dim cn As System.Data.OleDb.OleDbConnection
        Dim cmd As System.Data.OleDb.OleDbDataAdapter
        Dim ds As New System.Data.DataSet()

        cn = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;" & _
            "data source=" & PathNam & xlDateiName & ";Extended Properties=Excel 8.0;")

        ' Select the data from Sheet1 of the workbook.
        If xlwscnt > 1 Then
            cmd = New System.Data.OleDb.OleDbDataAdapter("select * from [Disc$]", cn)
        End If

        'cn.Open()
        'cmd.Fill(ds)
        'cn.Close()
    End Sub
    '##############################################################################################

    Private Sub TextBox6_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox6.TextChanged
        If Me.TextBox6.Text <> "" Then
            Me.TextBox6.BackColor = Color.White
        Else
            Me.TextBox6.BackColor = Color.LightSeaGreen
        End If
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        If Me.TextBox5.Text <> "" Then
            Me.TextBox5.BackColor = Color.White
        Else
            Me.TextBox5.BackColor = Color.LightSeaGreen
        End If
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged
        If Me.TextBox4.Text <> "" Then
            Me.TextBox4.BackColor = Color.White
        Else
            Me.TextBox4.BackColor = Color.LightSeaGreen
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

    Private Sub CheckBox2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox2.Click
        If Me.CheckBox2.Checked = True Then
            Me.TextBox4.Enabled = True
        Else
            Me.TextBox4.Enabled = False
        End If
    End Sub

    Private Sub ComboBox3_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.TextChanged
        If Me.ComboBox3.Text <> "" Then
            Me.ComboBox3.BackColor = Color.White
        Else
            Me.ComboBox3.BackColor = Color.LightSeaGreen
        End If
    End Sub
End Class