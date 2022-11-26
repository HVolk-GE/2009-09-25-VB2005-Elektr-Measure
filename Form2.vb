Option Strict Off
Option Explicit On
Friend Class Form2
	Inherits System.Windows.Forms.Form

    '================================================================
    ' LVCVTimer: 1 = CV ; 2 = LV
    ' DBOrExcel: 1 = DB; 2 = Excel 

    Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
        Dim i As Integer

        i = 55

        If Me.cmbAuswahl.Text = "" Then
            If State = "DE" Then
                MsgBox("Sie müssen einen Messzeitpunkt auswählen !", MsgBoxStyle.Critical, "Eintrag fehlt!")
            Else
                MsgBox("Select a time of measurement !", MsgBoxStyle.Critical, "Selection Failed !")
            End If
            Exit Sub
        End If

        If Me.RadioButton1.Checked = True Then
            If Me.RadioButton3.Checked = False And Me.RadioButton2.Checked = False Then
                If State = "DE" Then
                    MsgBox("Auswahl, Belag innen oder aussen fehlt !", MsgBoxStyle.Critical, "Auswahl fehlt")
                Else
                    MsgBox("Selection, pad inside or outside failed !", MsgBoxStyle.Critical, "Selection failed")
                End If
                Exit Sub
            End If

            If Me.RadioButton3.Checked = True Then
                If State = "DE" Then
                    Pads.Label6.Text = "Bremsbelage aussen"
                Else
                    Pads.Label6.Text = "Pad outside"
                End If
            End If

            If Me.RadioButton2.Checked = True Then
                If State = "DE" Then
                    Pads.Label6.Text = "Bremsbelage innen"
                Else
                    Pads.Label6.Text = "Pad inside"
                End If
            End If

            If selectmess = "" Then
                selectmess = Me.cmbAuswahl.Text
            Else
                Me.cmbAuswahl.Text = selectmess
                Me.cmbAuswahl.Enabled = False
            End If

        End If

        Pads.TextBox1.Text = Me.cmbAuswahl.Text
        Disc.TextBox1.Text = Me.cmbAuswahl.Text

        Disc.TextBox4.Text = DiscID
        Disc.TextBox5.Text = Runout
        Disc.ComboBox3.Text = DiscCondi

        If DiscID <> "" Then
            Disc.TextBox4.Enabled = False
        End If

        If DiscCondi <> "" Then
            Disc.ComboBox3.Enabled = False
        End If

        lngNumCols = CInt(Me.Label2.Text)

        Pads.Text1.Width = i
        Pads.Text2.Width = i
        Pads.Text3.Width = i
        Pads.Text4.Width = i
        Pads.Text5.Width = i
        Pads.Text6.Width = i
        Pads.Text7.Width = i
        Pads.Text8.Width = i
        Pads.Text9.Width = i
        Pads.Text10.Width = i
        Pads.Text11.Width = i
        Pads.Text12.Width = i

        canchelPad = 1

        If Padcnt > 0 Then
            If PadsDir = "updown" Then

                If lngNumCols = 4 Then
                    Pads.Text1.Location = New Point(8, 93)
                    Pads.Text2.Location = New Point(38, 189)

                    Pads.Text3.Location = New Point(430, 93)
                    Pads.Text4.Location = New Point(404, 189)

                ElseIf lngNumCols = 6 Then

                    Pads.Text1.Location = New Point(8, 93)
                    Pads.Text2.Location = New Point(38, 189)

                    Pads.Text3.Location = New Point(256, 24)
                    Pads.Text4.Location = New Point(256, 148)

                    Pads.Text5.Location = New Point(430, 93)
                    Pads.Text6.Location = New Point(404, 189)

                ElseIf lngNumCols = 8 Then

                    Pads.Text1.Location = New Point(8, 93)
                    Pads.Text2.Location = New Point(8, 189)

                    Pads.Text3.Location = New Point(188, 24)
                    Pads.Text4.Location = New Point(188, 148)

                    Pads.Text5.Location = New Point(256, 24)
                    Pads.Text6.Location = New Point(256, 148)

                    Pads.Text7.Location = New Point(430, 93)
                    Pads.Text8.Location = New Point(404, 189)

                ElseIf lngNumCols = 9 Then

                    Pads.Text1.Location = New Point(8, 93)
                    Pads.Text2.Location = New Point(8, 146)
                    Pads.Text3.Location = New Point(38, 189)

                    Pads.Text4.Location = New Point(256, 24)
                    Pads.Text5.Location = New Point(256, 77)
                    Pads.Text6.Location = New Point(256, 148)

                    Pads.Text7.Location = New Point(430, 93)
                    Pads.Text8.Location = New Point(430, 146)
                    Pads.Text9.Location = New Point(404, 189)

                ElseIf lngNumCols = 12 Then

                    Pads.Text1.Location = New Point(8, 93)
                    Pads.Text2.Location = New Point(8, 146)
                    Pads.Text3.Location = New Point(38, 189)

                    Pads.Text4.Location = New Point(188, 24)
                    Pads.Text5.Location = New Point(188, 77)
                    Pads.Text6.Location = New Point(188, 148)

                    Pads.Text7.Location = New Point(256, 24)
                    Pads.Text8.Location = New Point(256, 77)
                    Pads.Text9.Location = New Point(256, 148)

                    Pads.Text10.Location = New Point(430, 93)
                    Pads.Text11.Location = New Point(430, 146)
                    Pads.Text12.Location = New Point(404, 189)

                End If
            ElseIf PadsDir = "round" Then

                If lngNumCols = 4 Then

                    Pads.Text1.Location = New Point(8, 93)
                    Pads.Text2.Location = New Point(430, 93)

                    Pads.Text3.Location = New Point(38, 189)
                    Pads.Text4.Location = New Point(404, 189)

                ElseIf lngNumCols = 6 Then

                    Pads.Text1.Location = New Point(8, 93)
                    Pads.Text2.Location = New Point(256, 24)

                    Pads.Text3.Location = New Point(430, 93)
                    Pads.Text4.Location = New Point(38, 189)

                    Pads.Text5.Location = New Point(256, 148)
                    Pads.Text6.Location = New Point(404, 189)

                ElseIf lngNumCols = 8 Then

                    Pads.Text1.Location = New Point(8, 93)
                    Pads.Text2.Location = New Point(188, 24)

                    Pads.Text3.Location = New Point(256, 24)
                    Pads.Text4.Location = New Point(430, 93)

                    Pads.Text5.Location = New Point(8, 189)
                    Pads.Text6.Location = New Point(188, 148)

                    Pads.Text7.Location = New Point(256, 148)
                    Pads.Text8.Location = New Point(404, 189)

                ElseIf lngNumCols = 9 Then

                    Pads.Text1.Location = New Point(8, 93)
                    Pads.Text2.Location = New Point(256, 24)
                    Pads.Text3.Location = New Point(430, 93)

                    Pads.Text4.Location = New Point(8, 146)
                    Pads.Text5.Location = New Point(256, 77)
                    Pads.Text6.Location = New Point(430, 146)

                    Pads.Text7.Location = New Point(38, 189)
                    Pads.Text8.Location = New Point(256, 148)
                    Pads.Text9.Location = New Point(404, 189)

                ElseIf lngNumCols = 12 Then

                    Pads.Text1.Location = New Point(8, 93)
                    Pads.Text2.Location = New Point(188, 24)
                    Pads.Text3.Location = New Point(256, 24)

                    Pads.Text4.Location = New Point(430, 93)
                    Pads.Text5.Location = New Point(8, 146)
                    Pads.Text6.Location = New Point(188, 77)

                    Pads.Text7.Location = New Point(256, 77)
                    Pads.Text8.Location = New Point(430, 146)
                    Pads.Text9.Location = New Point(38, 189)

                    Pads.Text10.Location = New Point(188, 148)
                    Pads.Text11.Location = New Point(256, 148)
                    Pads.Text12.Location = New Point(404, 189)

                End If
            End If

            If lngNumCols = 4 Then

                Pads.Text1.Visible = True
                Pads.Text1.Enabled = True

                Pads.Text2.Visible = True
                Pads.Text2.Enabled = True

                Pads.Text3.Visible = True
                Pads.Text3.Enabled = True

                Pads.Text4.Visible = True
                Pads.Text4.Enabled = True

                Pads.Text5.Visible = False
                Pads.Text5.Enabled = False
                Pads.Text6.Visible = False
                Pads.Text6.Enabled = False
                Pads.Text7.Visible = False
                Pads.Text7.Enabled = False
                Pads.Text8.Visible = False
                Pads.Text8.Enabled = False
                Pads.Text9.Visible = False
                Pads.Text9.Enabled = False
                Pads.Text10.Visible = False
                Pads.Text10.Enabled = False
                Pads.Text11.Visible = False
                Pads.Text11.Enabled = False
                Pads.Text12.Visible = False
                Pads.Text12.Enabled = False
                Pads.Show()
                Me.Hide()

            ElseIf lngNumCols = 6 Then

                Pads.Text1.Visible = True
                Pads.Text1.Enabled = True
                Pads.Text2.Visible = True
                Pads.Text2.Enabled = True
                Pads.Text3.Visible = True
                Pads.Text3.Enabled = True
                Pads.Text4.Visible = True
                Pads.Text4.Enabled = True
                Pads.Text5.Visible = True
                Pads.Text5.Enabled = True
                Pads.Text6.Visible = True
                Pads.Text6.Enabled = True

                Pads.Text7.Visible = False
                Pads.Text7.Enabled = False
                Pads.Text8.Visible = False
                Pads.Text8.Enabled = False
                Pads.Text9.Visible = False
                Pads.Text9.Enabled = False
                Pads.Text9.Visible = False
                Pads.Text9.Enabled = False
                Pads.Text10.Visible = False
                Pads.Text10.Enabled = False
                Pads.Text11.Visible = False
                Pads.Text11.Enabled = False
                Pads.Text12.Visible = False
                Pads.Text12.Enabled = False

                Pads.Show()
                Me.Hide()

            ElseIf lngNumCols = 8 Then

                Pads.Text1.Visible = True
                Pads.Text1.Enabled = True
                Pads.Text2.Visible = True
                Pads.Text2.Enabled = True
                Pads.Text3.Visible = True
                Pads.Text3.Enabled = True
                Pads.Text4.Visible = True
                Pads.Text4.Enabled = True
                Pads.Text5.Visible = True
                Pads.Text5.Enabled = True
                Pads.Text6.Visible = True
                Pads.Text6.Enabled = True
                Pads.Text7.Visible = True
                Pads.Text7.Enabled = True
                Pads.Text8.Visible = True
                Pads.Text8.Enabled = True

                Pads.Text9.Visible = False
                Pads.Text9.Enabled = False
                Pads.Text10.Visible = False
                Pads.Text10.Enabled = False
                Pads.Text11.Visible = False
                Pads.Text11.Enabled = False
                Pads.Text12.Visible = False
                Pads.Text12.Enabled = False

                Pads.Show()
                Me.Hide()

            ElseIf lngNumCols = 9 Then

                Pads.Text1.Visible = True
                Pads.Text1.Enabled = True
                Pads.Text2.Visible = True
                Pads.Text2.Enabled = True
                Pads.Text3.Visible = True
                Pads.Text3.Enabled = True
                Pads.Text4.Visible = True
                Pads.Text4.Enabled = True
                Pads.Text5.Visible = True
                Pads.Text5.Enabled = True
                Pads.Text6.Visible = True
                Pads.Text6.Enabled = True
                Pads.Text7.Visible = True
                Pads.Text7.Enabled = True
                Pads.Text8.Visible = True
                Pads.Text8.Enabled = True
                Pads.Text9.Visible = True
                Pads.Text9.Enabled = True

                Pads.Text10.Visible = False
                Pads.Text10.Enabled = False
                Pads.Text11.Visible = False
                Pads.Text11.Enabled = False
                Pads.Text12.Visible = False
                Pads.Text12.Enabled = False

                Pads.Show()
                Me.Hide()
            ElseIf lngNumCols = 12 Then

                Pads.Text1.Visible = True
                Pads.Text1.Enabled = True
                Pads.Text2.Visible = True
                Pads.Text2.Enabled = True
                Pads.Text3.Visible = True
                Pads.Text3.Enabled = True
                Pads.Text4.Visible = True
                Pads.Text4.Enabled = True
                Pads.Text5.Visible = True
                Pads.Text5.Enabled = True
                Pads.Text6.Visible = True
                Pads.Text6.Enabled = True
                Pads.Text7.Visible = True
                Pads.Text7.Enabled = True
                Pads.Text8.Visible = True
                Pads.Text8.Enabled = True
                Pads.Text9.Visible = True
                Pads.Text9.Enabled = True
                Pads.Text9.Visible = True
                Pads.Text9.Enabled = True
                Pads.Text10.Visible = True
                Pads.Text10.Enabled = True
                Pads.Text11.Visible = True
                Pads.Text11.Enabled = True
                Pads.Text12.Visible = True
                Pads.Text12.Enabled = True

                Pads.Show()
                Me.Hide()
            End If

        ElseIf Disccnt > 0 Then

            If Microsoft.VisualBasic.Left(xlDateiName, 1) = StartLVTestNr _
                Or Microsoft.VisualBasic.Left(DBTestnr, 1) = StartLVTestNr Then

                Disc.Text1.Location = New Point(248, 69)
                Disc.Text2.Location = New Point(130, 107)
                Disc.Text3.Location = New Point(49, 193)
                Disc.Text4.Location = New Point(16, 303)
                Disc.Text5.Location = New Point(49, 434)
                Disc.Text6.Location = New Point(130, 512)
                Disc.Text7.Location = New Point(248, 553)
                Disc.Text8.Location = New Point(377, 512)
                Disc.Text9.Location = New Point(446, 434)
                Disc.Text10.Location = New Point(480, 303)
                Disc.Text11.Location = New Point(446, 193)
                Disc.Text12.Location = New Point(337, 107)

                Disc.Label17.Visible = True
                Disc.Label16.Visible = True
                Disc.Label15.Visible = True
                Disc.Label14.Visible = True
                Disc.Label13.Visible = True
                Disc.Label6.Visible = True
                Disc.Label7.Visible = True
                Disc.Label12.Visible = True

                Disc.Label20.Visible = True
                Disc.ComboBox3.Visible = True

                Disc.Label17.Location = New Point(245, 92)
                Disc.Label16.Location = New Point(131, 130)
                Disc.Label15.Location = New Point(44, 216)
                Disc.Label14.Location = New Point(22, 323)
                Disc.Label13.Location = New Point(46, 417)
                Disc.Label12.Location = New Point(372, 130)
                Disc.Label11.Location = New Point(443, 216)
                Disc.Label10.Location = New Point(475, 326)
                Disc.Label9.Location = New Point(447, 418)
                Disc.Label8.Location = New Point(378, 496)
                Disc.Label7.Location = New Point(245, 537)
                Disc.Label6.Location = New Point(131, 496)

                If lngNumCols = 4 Then

                    Disc.Text2.Location = Disc.Text4.Location
                    Disc.Text3.Location = Disc.Text7.Location
                    Disc.Text4.Location = Disc.Text10.Location

                    Disc.Text5.Visible = False
                    Disc.Text6.Visible = False
                    Disc.Text7.Visible = False
                    Disc.Text8.Visible = False
                    Disc.Text9.Visible = False
                    Disc.Text10.Visible = False
                    Disc.Text11.Visible = False
                    Disc.Text12.Visible = False

                    Disc.Text5.Enabled = False
                    Disc.Text6.Enabled = False
                    Disc.Text7.Enabled = False
                    Disc.Text8.Enabled = False
                    Disc.Text9.Enabled = False
                    Disc.Text10.Enabled = False
                    Disc.Text11.Enabled = False
                    Disc.Text12.Enabled = False

                    Disc.Label17.Text = "1"
                    Disc.Label16.Visible = False
                    Disc.Label15.Visible = False
                    Disc.Label14.Text = "4"
                    Disc.Label13.Visible = False
                    Disc.Label6.Visible = False
                    Disc.Label7.Text = "7"
                    Disc.Label8.Visible = False
                    Disc.Label9.Visible = False
                    Disc.Label10.Text = "10"
                    Disc.Label11.Visible = False
                    Disc.Label12.Visible = False

                ElseIf lngNumCols > 4 Then
                    Disc.Label17.Text = "1"
                    Disc.Label16.Text = "2"
                    Disc.Label15.Text = "3"
                    Disc.Label14.Text = "4"
                    Disc.Label13.Text = "5"
                    Disc.Label6.Text = "6"
                    Disc.Label7.Text = "7"
                    Disc.Label8.Text = "8"
                    Disc.Label9.Text = "9"
                    Disc.Label10.Text = "10"
                    Disc.Label11.Text = "11"
                    Disc.Label12.Text = "12"
                End If
                Disc.Show()
                Me.Hide()
            Else
                Disc.Show()
                Me.Hide()
            End If

        End If

    End Sub

    Private Sub Form2_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim ctl As System.Windows.Forms.Control
        Dim savetrue As Integer

        canchelPad = 0

        If savetrue = 0 Then
            savetrue = 1
        End If

        selectmess = ""

        Me.ToolStripStatusLabel1.Text = Creatxt

        If frmLogin.Check1.Checked = True Then
            Padcnt = 1
            Disccnt = 0
            PadInside = 0
            Me.RadioButton1.Checked = True
        Else
            Me.RadioButton1.Checked = False
        End If

        If frmLogin.Check2.Checked = True Then
            Padcnt = 1
            Disccnt = 0
            PadInside = 1
            Padoutside = 0
            Me.RadioButton2.Checked = True
        Else
            Me.RadioButton2.Checked = False
        End If

        If frmLogin.Check3.Checked = True Then
            Padcnt = 1
            Disccnt = 0
            Padoutside = 1
            PadInside = 0
            Me.RadioButton3.Checked = True
        Else
            Me.RadioButton3.Checked = False
        End If

        If frmLogin.Check4.Checked = True Then
            Padcnt = 0
            Disccnt = 1
            Padoutside = 0
            PadInside = 0
            Me.RadioButton4.Checked = True
        Else
            Me.RadioButton4.Checked = False
        End If

        If State = "DE" Then
            Me.Label1.Text = "Messpunkte : "
            Me.Label3.Text = "Kontrolle"
            Me.Command1.Text = "OK"
            Me.Command2.Text = "Schliessen"
            Me.RadioButton1.Text = "Belag"
            Me.RadioButton2.Text = "Innen"
            Me.RadioButton3.Text = "Aussen"
            Me.RadioButton4.Text = "Scheibe"
        Else
            Me.Label1.Text = "Measure points : "
            Me.Label3.Text = "Control"
            Me.Command1.Text = "OK"
            Me.Command2.Text = "Close"
            Me.RadioButton1.Text = "Pads"
            Me.RadioButton2.Text = "Inside"
            Me.RadioButton3.Text = "Outside"
            Me.RadioButton4.Text = "Disc"
        End If

        If DBOrExcel = 2 Then
            ' LVCVTimer: 1 = CV ; 2 = LV
            If LVCVTimer = 1 Then
                readExcel()
            ElseIf LVCVTimer = 2 Then
                LVMircofaceExcel()
            End If
        End If

        lngNumCols = CInt(Me.Label2.Text)

        Me.cmbAuswahl.Enabled = True
        Me.Command2.Enabled = True

        Me.RadioButton1.Checked = True
        Me.RadioButton2.Checked = True

        If Me.cmbAuswahl.Items.Count = 0 And savetrue = 1 Then
            CheckcmdAuswahl()
        End If

        Me.TopMost = True

        Exit Sub

err_Handler:
        'Wenn ein Fehler aufgetreten ist, alle Controls
        '  ausser Beenden-Button deaktivieren:

        For Each ctl In Me.Controls
            If ctl.Name <> "Command2" Then
                ctl.Enabled = False
            End If
        Next ctl
        Exit Sub
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        Dim i As Integer

        If DBOrExcel = 2 Then

            Padcnt = 1
            Disccnt = 0

            Me.RadioButton2.Enabled = True
            Me.RadioButton3.Enabled = True

            CancelChecks()

        ElseIf DBOrExcel = 1 Then

            Padcnt = 1
            Disccnt = 0

            If Me.RadioButton1.Checked = True Then
                Me.cmbAuswahl.Items.Clear()

                Me.RadioButton2.Enabled = True
                Me.RadioButton3.Enabled = True

                Me.Label2.Text = CStr(MesurementPointPads)

                PadDisc = "Pad"

                If MeasurmentTimesPads = 2 Then
                    Me.cmbAuswahl.Items.Add("Start")
                    Me.cmbAuswahl.Items.Add("End")
                End If

                If MeasurmentTimesPads > 2 Then
                    For i = 1 To MeasurmentTimesPads
                        If i = 1 Then
                            Me.cmbAuswahl.Items.Add("Start")
                        ElseIf i = MeasurmentTimesPads Then
                            Me.cmbAuswahl.Items.Add("End")
                        Else
                            Me.cmbAuswahl.Items.Add(i - 1 & "ZM")
                        End If
                    Next
                End If
            End If
        End If

    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged

        If DBOrExcel = 2 Then

            PadInside = 1
            Padoutside = 0
            Sheets01 = ""
            Sheets02 = ""
            Sheets03 = ""

            CancelChecks()

            If selectmess = "" Then
                selectmess = Me.cmbAuswahl.Text
            Else
                Me.cmbAuswahl.Text = selectmess
                Me.cmbAuswahl.Enabled = False
            End If

        ElseIf DBOrExcel = 1 Then

            If Me.RadioButton2.Checked = True Then

                PadInside = 1
                Padoutside = 0
                Sheets01 = ""
                Sheets02 = ""
                Sheets03 = ""
                If selectmess = "" Then
                    selectmess = Me.cmbAuswahl.Text
                Else
                    Me.cmbAuswahl.Text = selectmess
                    Me.cmbAuswahl.Enabled = False
                End If
            End If

        End If

    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged

        If DBOrExcel = 2 Then

            Padoutside = 1
            PadInside = 0
            Sheets01 = ""
            Sheets02 = ""
            Sheets03 = ""

            CancelChecks()

            If selectmess = "" Then
                selectmess = Me.cmbAuswahl.Text
            Else
                Me.cmbAuswahl.Text = selectmess
                Me.cmbAuswahl.Enabled = False
            End If

        ElseIf DBOrExcel = 1 Then

            Padoutside = 1
            PadInside = 0
            Sheets01 = ""
            Sheets02 = ""
            Sheets03 = ""

            If selectmess = "" Then
                selectmess = Me.cmbAuswahl.Text
            Else
                Me.cmbAuswahl.Text = selectmess
                Me.cmbAuswahl.Enabled = False
            End If

        End If

    End Sub

    Private Sub RadioButton4_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton4.CheckedChanged
        Dim i As Integer

        If DBOrExcel = 2 Then

            Disccnt = 1
            Padoutside = 0
            PadInside = 0
            Padcnt = 0

            If Me.RadioButton4.Checked = True Then


                Me.RadioButton2.Enabled = False
                Me.RadioButton3.Enabled = False
                Me.RadioButton2.Checked = False
                Me.RadioButton3.Checked = False

                CancelChecks()

                Me.cmbAuswahl.Enabled = True
            End If

        ElseIf DBOrExcel = 1 Then
            Disccnt = 1
            Padoutside = 0
            PadInside = 0
            Padcnt = 0

            If Me.RadioButton4.Checked = True Then

                Me.cmbAuswahl.Items.Clear()

                Me.RadioButton3.Enabled = False
                Me.RadioButton2.Enabled = False

                Me.Label2.Text = CStr(MesurementPointDisc)

                PadDisc = "Disc"

                If MeasurmentTimesDisc = 2 Then
                    Me.cmbAuswahl.Items.Add("Start")
                    Me.cmbAuswahl.Items.Add("End")
                End If

                If MeasurmentTimesDisc > 2 Then
                    For i = 1 To MeasurmentTimesDisc
                        If i = 1 Then
                            Me.cmbAuswahl.Items.Add("Start")
                        ElseIf i = MeasurmentTimesDisc Then
                            Me.cmbAuswahl.Items.Add("End")
                        Else
                            Me.cmbAuswahl.Items.Add(i - 1 & "ZM")
                        End If
                    Next
                End If

                Me.RadioButton2.Enabled = False
                Me.RadioButton3.Enabled = False
                Me.RadioButton2.Checked = False
                Me.RadioButton3.Checked = False

                Me.cmbAuswahl.Enabled = True

            End If
        End If
    End Sub

    Private Sub cmbAuswahl_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbAuswahl.SelectedIndexChanged
        'Werte aus Excel-Zellen in TextBox 1 - 5 einlesen:
        Dim lngRowIndex As Integer
        Dim intColIndex As Short

        If DBOrExcel = 2 Then

            lngRowIndex = cmbAuswahl.SelectedIndex + 1

            ' LVCVTimer: 1 = CV ; 2 = LV
            If LVCVTimer = 1 Then
                lngNumCols = xlWS.Range("A1").End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight).Column
                For intColIndex = 1 To lngNumCols - 1
                    If Me.cmbAuswahl.Text = xlWS.Cells(lngRowIndex, 1).Value Then
                        xlWS.Cells(lngRowIndex, 2).Select()
                        startWriteRow = lngRowIndex
                        startWriteColum = 2
                        If Disccnt = 1 Then
                            DiscID = xlWS.Cells(2, lngTotalNumCols - 2).Value
                            If StartLVTestNr = Microsoft.VisualBasic.Left(testnumber, 1) Then
                                DiscCondi = xlWS.Cells(40, 14).Value
                            End If
                        End If
                        lngNumCols = lngTotalNumCols - 3
                        cntStartRow = lngRowIndex
                        Exit For
                    End If
                    lngRowIndex = lngRowIndex + 1
                Next intColIndex
            ElseIf LVCVTimer = 2 And tttemp <> "" Then
                lngNumCols = xlWS.Range("Z" & txtCnt).End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight).Column
                For intColIndex = intColIndex0 To intColIndex0 + MessCnt
                    If txtCnt > lngRowIndex Then lngRowIndex = txtCnt + 1
                    If Me.cmbAuswahl.Text = xlWS.Cells(lngRowIndex, intColIndex0).Value Then
                        xlWS.Cells(lngRowIndex, intColIndex0 + 1).Select()
                        startWriteRow = lngRowIndex
                        startWriteColum = intColIndex0 + 1
                        If Disccnt = 1 Then
                            DiscID = xlWS.Cells(txtCnt + 1, startWriteColum + lngTotalNumCols - 3).Value
                            DiscCondi = xlWS.Cells(40, 14).Value
                        End If
                        lngNumCols = lngTotalNumCols - 3
                        cntStartRow = lngRowIndex
                        Exit For
                    End If
                    lngRowIndex = lngRowIndex + 1
                Next intColIndex

            ElseIf LVCVTimer = 2 And tttemp = "" Then

                lngNumCols = xlWS.Range("A1").End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight).Column
                For intColIndex = 1 To lngNumCols - 1 '5
                    If Me.cmbAuswahl.Text = xlWS.Cells(lngRowIndex, 1).Value Then
                        xlWS.Cells(lngRowIndex, 2).Select()
                        startWriteRow = lngRowIndex
                        startWriteColum = 2
                        If Disccnt = 1 Then
                            DiscID = xlWS.Cells(2, lngTotalNumCols - 2).Value
                            DiscCondi = xlWS.Cells(40, 14).Value
                        End If
                        lngNumCols = lngTotalNumCols - 3
                        cntStartRow = lngRowIndex
                        Exit For
                    End If
                    lngRowIndex = lngRowIndex + 1
                Next intColIndex
            End If

        End If

    End Sub
	
    Public Sub Form2_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        'Excel "aufräumen":
        On Error Resume Next
        If DBOrExcel = 2 Then
            If Not xlAppl Is Nothing Then
                If Not xlWB Is Nothing Then
                    'Verweise freigeben:
                    xlWS = Nothing
                    xlWB = Nothing
                    wb = Nothing
                    boolWBOffen = False
                End If

                'Wenn Excel nicht bereits ausgeführt wurde, schliessen:
                If xlApplLiefNicht Then xlAppl.Application.Quit()
                'Verweis freigeben:
                boolWBOffen = False
                xlAppl = Nothing
            End If

        End If

        eventArgs.Cancel = Cancel

    End Sub
	
    Private Sub Command2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Command2.Click

        If DBOrExcel = 2 Then
            frmLogin.Option1.Enabled = False
            frmLogin.Option2.Enabled = False
            frmLogin.Check1.Enabled = False
            frmLogin.Check2.Enabled = False
            frmLogin.Check3.Enabled = False
            frmLogin.Check4.Enabled = False

            frmLogin.Option1.Checked = False
            frmLogin.Option2.Checked = False
            frmLogin.Check1.Checked = False
            frmLogin.Check2.Checked = False
            frmLogin.Check3.Checked = False
            frmLogin.Check4.Checked = False

            frmLogin.Combo1.Text = ""
            frmLogin.Combo1.Items.Clear()
            frmLogin.File1.Items.Clear()

            frmLogin.txtUserName.Text = ""

            Me.TopMost = False
            Me.Hide()

            If xlAppl.ActiveWorkbook.Sheets.Count = 1 Then
                xlAppl.DisplayAlerts = False
                xlAppl.ActiveWorkbook.SaveAs(Filename:=PathNam & xlDateiName, FileFormat:= _
                            Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel4, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False _
                            , CreateBackup:=False)
                xlWB.Close()
                xlAppl.DisplayAlerts = True
            ElseIf xlAppl.ActiveWorkbook.Sheets.Count > 1 Then
                xlAppl.DisplayAlerts = False
                xlWB.Close(SaveChanges:=True)
                xlAppl.DisplayAlerts = True
            End If


            Me.Close()
            frmLogin.Show()

        ElseIf DBOrExcel = 1 Then

            frmLoginDB.Label3.Visible = False
            frmLoginDB.ComboBox2.Visible = False
            frmLoginDB.ComboBox1.Enabled = False
            frmLoginDB.CheckBox1.Enabled = False
            cntBack0 = 2
            Me.Close()
            frmLoginDB.Show()

        End If

    End Sub

    Sub CancelChecks()
        Dim i As Integer

        If DBOrExcel = 2 Then

            If LVCVTimer = 1 Then
                ' 0 = Keine Auswahl, 1 = Auswahl
                If Disccnt > 0 And Padcnt < 1 Then
                    LastProj1 = GetIniString("Sheets", "CVSheets01", INIPath) '***
                    Sheets01 = LastProj1
                ElseIf Padcnt > 0 And PadInside > 0 Then
                    LastProj1 = GetIniString("Sheets", "CVSheets02", INIPath) '***
                    Sheets02 = LastProj1
                ElseIf Padcnt > 0 And Padoutside > 0 Then
                    LastProj1 = GetIniString("Sheets", "CVSheets03", INIPath) '***
                    Sheets03 = LastProj1
                End If

            ElseIf LVCVTimer = 2 Then

                If MFCnt = 1 Then

                    If tttemp <> "" Then

                        If PrgCnt >= 1 Then

                            For i = 1 To PrgCnt
                                If tttemp = Prg(i) Then
                                    If Disccnt > 0 And Padcnt < 1 Then
                                        LastProj1 = GetIniString(tttemp, "Disc", INIPath) '***
                                        Sheets01 = LastProj1
                                        LastProj1 = GetIniString(tttemp, "DiscMeasurePoints", INIPath) '***
                                        MessCnt = LastProj1
                                    ElseIf Padcnt > 0 And PadInside > 0 Then
                                        LastProj1 = GetIniString(tttemp, "PadInSide", INIPath) '***
                                        Sheets02 = LastProj1
                                        LastProj1 = GetIniString(tttemp, "PadMeasurePoints", INIPath) '***
                                        MessCnt = LastProj1
                                    ElseIf Padcnt > 0 And Padoutside > 0 Then
                                        LastProj1 = GetIniString(tttemp, "PadOutSide", INIPath) '***
                                        Sheets03 = LastProj1
                                        LastProj1 = GetIniString(tttemp, "PadMeasurePoints", INIPath) '***
                                        MessCnt = LastProj1
                                    End If
                                End If
                            Next
                        End If
                    Else
                        If Disccnt > 0 And Padcnt < 1 Then
                            LastProj1 = GetIniString("Sheets", "LVSheets01", INIPath) '***
                            Sheets01 = LastProj1
                        ElseIf Padcnt > 0 And PadInside > 0 Then
                            LastProj1 = GetIniString("Sheets", "LVSheets02", INIPath) '***
                            Sheets02 = LastProj1
                        ElseIf Padcnt > 0 And Padoutside > 0 Then
                            LastProj1 = GetIniString("Sheets", "LVSheets03", INIPath) '***
                            Sheets03 = LastProj1
                        End If
                    End If
                End If
            ElseIf MFCnt = 0 Then
                If Disccnt > 0 And Padcnt < 1 Then
                    LastProj1 = GetIniString("Sheets", "LVSheets01", INIPath) '***
                    Sheets01 = LastProj1
                ElseIf Padcnt > 0 And PadInside > 0 Then
                    LastProj1 = GetIniString("Sheets", "LVSheets02", INIPath) '***
                    Sheets02 = LastProj1
                ElseIf Padcnt > 0 And Padoutside > 0 Then
                    LastProj1 = GetIniString("Sheets", "LVSheets03", INIPath) '***
                    Sheets03 = LastProj1
                End If
            End If

            If LVCVTimer = 1 Then
                readExcel()
            ElseIf LVCVTimer = 2 Then
                LVMircofaceExcel()
            End If

        End If

    End Sub

    Private Sub Form2_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.VisibleChanged

        If canchelPad = 0 Then
            Me.cmbAuswahl.Text = ""
            Me.cmbAuswahl.Refresh()
            Me.cmbAuswahl.Enabled = True
            selectmess = ""
            Exit Sub
        End If

        If selectmess = "" Then
            selectmess = Me.cmbAuswahl.Text
            CancelChecks()
        Else
            Me.cmbAuswahl.Text = selectmess
            Me.cmbAuswahl.Enabled = False
            CancelChecks()
        End If

    End Sub

End Class
