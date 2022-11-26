Public Class MainMenue

    Private Sub MainMenue_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        IniTal()

        If State = "DE" Then
            Me.Text = "Hauptmenu"
            Me.Testeintrag1ToolStripMenuItem.Text = "Messungen"
            Me.VerschleissmessungenToolStripMenuItem.Text = "Verschleißmessungen"
            Me.ListeVonAktuellenTestsToolStripMenuItem.Text = "Liste von aktuellen Tests"
            Me.InformationenToolStripMenuItem.Text = "Informationen"
            Me.ViewWearTemplatesToolStripMenuItem.Text = "View Templates"
            Me.DTVMessungToolStripMenuItem.Text = "DTV Messung"
            Me.PhotografieVorlagenToolStripMenuItem.Text = "Photografie Vorlagen"
            Me.KonfigToolStripMenuItem.Text = "Konfiguration"
            Me.AnwendungKonfigurierenToolStripMenuItem.Text = "Anwendung konfigurieren"
            Me.HilfeToolStripMenuItem.Text = "Hilfe"
            Me.HilfeToolStripMenuItem1.Text = "Dokumentation"
            Me.KurzAnleitungToolStripMenuItem.Text = "Kurz Anleitung"
            Me.BeendenToolStripMenuItem.Text = "Beenden"
            Me.Button1.Text = "Schliess Listen Ansicht"
            Me.CheckBox1.Text = "Öffne die ausgewählte Datei"
            Me.Label2.Text = "Messung - Tools, dieses Tool stellt eine " & Chr(10) & _
                             "Arbeitserleichterung dar." & Chr(10) & _
                             "In diesem Tool wird zur Zeit, das Messen " & Chr(10) & _
                             "mit Bügelmeßschraube und das wiegen mit " & Chr(10) & _
                             "der Kern - Waage unterstützt." 
        Else
            Me.Text = "Main Menue"
            Me.Testeintrag1ToolStripMenuItem.Text = "Measurement"
            Me.VerschleissmessungenToolStripMenuItem.Text = "Wear Measurement"
            Me.ListeVonAktuellenTestsToolStripMenuItem.Text = "List of available Tests"
            Me.InformationenToolStripMenuItem.Text = "Informationen"
            Me.DTVMessungToolStripMenuItem.Text = "DTV Scan"
            Me.ViewWearTemplatesToolStripMenuItem.Text = "View Templates"
            Me.PhotografieVorlagenToolStripMenuItem.Text = "Photograph Samples"
            Me.KonfigToolStripMenuItem.Text = "Config"
            Me.AnwendungKonfigurierenToolStripMenuItem.Text = "Application Config"
            Me.HilfeToolStripMenuItem.Text = "Help"
            Me.HilfeToolStripMenuItem1.Text = "Documentation (de)"
            Me.KurzAnleitungToolStripMenuItem.Text = "Short documenation (de)"
            Me.BeendenToolStripMenuItem.Text = "Quit"
            Me.Button1.Text = "Quit List view"
            Me.CheckBox1.Text = "Open select file"
            Me.Label2.Text = "Measurement - Tools, these tool is in" & Chr(10) & _
                             "to time a working help." & Chr(10) & _
                             "In this  time, can you use" & Chr(10) & _
                             "for Measurement and weigh" & Chr(10) & _
                             "Micro - Mitutoyo and" & Chr(10) & _
                             "Kern - weigh machine"
        End If

        Me.File1.Visible = False
        Me.Button1.Visible = False
        Me.Label1.Visible = False
        Me.CheckBox1.Visible = False
        Me.CheckBox1.Checked = False

        Me.PhotografieVorlagenToolStripMenuItem.Enabled = False
        Me.DTVMessungToolStripMenuItem.Enabled = False
        '* Me.HilfeToolStripMenuItem.Enabled = False
        'Me.HilfeToolStripMenuItem.Text = "Help"
        'Me.HilfeToolStripMenuItem1.Text = "Help"
        Me.KurzAnleitungToolStripMenuItem.Enabled = False
        Me.COMPortTestToolStripMenuItem.Visible = False
        Me.StepByStepToolStripMenuItem.Visible = False

        Me.Label3.Text = "Federal Mogul." & Chr(10) & _
                         "Direct PC Measurement " & Chr(10) & _
                         "© by H. Volk" & Chr(10)

        Version.Text = System.String.Format(Version.Text, My.Application.Info.Version.Major, My.Application.Info.Version.Minor)

        Me.ViewWearTemplatesToolStripMenuItem.Enabled = False
        
        Me.TopMost = True

    End Sub

    Private Sub BeendenToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BeendenToolStripMenuItem.Click
        End
    End Sub

    Private Sub VerschleissmessungenToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles VerschleissmessungenToolStripMenuItem.Click
        If DBOrExcel = 2 Then
            frmLogin.Show()
        ElseIf DBOrExcel = 1 Then
            frmLoginDB.Show()
        ElseIf DBOrExcel = 0 Then
            If State = "DE" Then
                MsgBox("Fehler, es ist kein Ausgabe Format konfiguriert !", MsgBoxStyle.Critical, "Fehler !")
            Else
                MsgBox("Failed, not output Format config !", MsgBoxStyle.Critical, "Failed !")
            End If
            Exit Sub
        End If
        Me.Hide()
    End Sub

    Private Sub ListeVonAktuellenTestsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListeVonAktuellenTestsToolStripMenuItem.Click
        Me.File1.Path = CVPath
        Me.Label1.Text = CVPath
        Me.File1.FileName = CVFile
        Me.File1.Refresh()
        Me.Label2.Visible = False
        Me.File1.Visible = True
        Me.Button1.Visible = True
        Me.Label1.Visible = True
        Me.CheckBox1.Visible = True
    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.File1.Visible = False
        Me.Button1.Visible = False
        Me.Label1.Visible = False
        Me.CheckBox1.Visible = False
        Me.Label2.Visible = True
    End Sub

    Private Sub File1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles File1.SelectedValueChanged
        Dim strFilenam As String

        strFilenam = ""
        strFilenam = Me.File1.SelectedItem.ToString

        xlDateiName = strFilenam

    End Sub

    Sub OpenInExcel()
        '*Dim boolWBOffen As Boolean
        Dim i As Short
        '*Dim wb As Object ' As Excel.Workbook
        '*Dim wb As Microsoft.Office.Interop.Excel.Workbook

        ' GetObject-Funktionsaufruf ohne erstes Argument gibt einen
        ' Verweis auf eine Instanz der Anwendung zurück. Wenn die
        ' Anwendung nicht ausgeführt wird, tritt ein Fehler auf.

        'Prüfen, ob Excel ausgeführt wird:
        On Error Resume Next
        '        xlAppl = GetObject(, "Excel.Application")
        '*        xlAppl = CType(CreateObject("Excel.Application"), _
        '*                             Microsoft.Office.Interop.Excel.Application)

        If Err.Number <> 0 Then xlApplLiefNicht = True

        Err.Clear() ' Err-Objekt im Fehlerfall löschen

        'Wenn Excel nicht ausgeführt wird, Excel starten:
        If xlAppl Is Nothing Then
            '            xlAppl = CreateObject("Excel.Application")
            xlAppl = CType(CreateObject("Excel.Application"), _
                                   Microsoft.Office.Interop.Excel.Application)

            'Wenn ein Fehler aufgetreten ist...
            If Err.Number <> 0 Then
                MsgBox("Konnte keine Verbindung zu Excel herstellen !", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, Form2.Text)
                GoTo err_Handler
            End If
        End If

        'Prüfen, ob Arbeitsmappe bereits offen ist:
        '*boolWBOffen = False

        If Not xlApplLiefNicht Then
            If xlAppl.Workbooks.Count > 0 Then
                For Each wb In xlAppl.Workbooks
                    If LCase(wb.Name) = LCase(xlDateiName) Then
                        'wb.Activate
                        boolWBOffen = True
                        Exit For
                    End If
                Next wb
            End If
        End If

        'LVCVTimer: 1 = CV ; 2 = LV
        If frmLogin.Option1.Visible = True Then
            If LVCVTimer = 1 Then
                PathNam = CVPath
            ElseIf LVCVTimer = 2 Then
                PathNam = LVPath
            End If
        Else
            PathNam = CVPath
        End If

        On Error Resume Next
        If Not boolWBOffen Then
            'Wenn Arbeitsmappe nicht offen, öffnen und Verweis setzen:
            Err.Clear()
            xlWB = xlAppl.Workbooks.Open(FileName:=PathNam & xlDateiName)

            'Wenn ein Fehler aufgetreten ist...
            If Err.Number <> 0 Then
                MsgBox("Die Arbeitsmappe '" & xlDateiName & "' konnte nicht geöffnet werden !", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, Form2.Text)
                If xlApplLiefNicht Then xlAppl.Application.Quit()
                xlAppl = Nothing
                GoTo err_Handler
            End If

        Else
            'Verweis setzen:
            xlWB = xlAppl.Workbooks(xlDateiName)
        End If

        On Error GoTo 0

        If excelview = "True" Then
            xlAppl.Application.Visible = True
        Else
            xlAppl.Application.Visible = False
        End If

        '*xlAppl.Application.Visible = True

        xlDateiName = ""

        Exit Sub

err_Handler:
        'Wenn ein Fehler aufgetreten ist, alle Controls
        'ausser Beenden-Button deaktivieren:
        Dim ctl As System.Windows.Forms.Control

        '*For Each ctl In Form2.Controls
        '*If ctl.Name <> "cmdBeenden" Then
        '*ctl.Enabled = False
        '*End If
        '*Next ctl

        Exit Sub

    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If Me.CheckBox1.Checked = True Then
            OpenInExcel()
        End If
    End Sub

    Private Sub COMPortTestToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles COMPortTestToolStripMenuItem.Click
        Me.Hide()
        'Me.Close()
        NETSerialTerm.Show()
    End Sub

    Private Sub StepByStepToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles StepByStepToolStripMenuItem.Click
        'Dim Pathstr00 As String
        '("Excel.Application"), _
        '        Microsoft.Office.Interop.Excel.Application)
        'Pathstr00 = "C:\Dokumente und Einstellungen\volkh\Eigene Dateien\Visual Studio 2005\Projects\Elektr-Measure\Projekt1.NET\Documentation\PC Direkt Elektr Step by Step.htm"
        'Shell(Pathstr00)
        'System.IO.File.Open(Pathstr00)
    End Sub

    Private Sub AnwendungKonfigurierenToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles AnwendungKonfigurierenToolStripMenuItem.Click

        Me.Hide()
        '.Close()
        ApplConfig.Show()

    End Sub

    Private Sub DTVMessungToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DTVMessungToolStripMenuItem.Click
        Dim DTVStart As String
        DTVStart = DTVPath & DTVAnw

        Me.TopMost = False

        Process.Start(DTVStart)

    End Sub

    Private Sub HilfeToolStripMenuItem1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles HilfeToolStripMenuItem1.Click
        Dim Help1Start As String
        '
        'If State = "DE" Then
        'Help1Start = My.Application.Info.DirectoryPath & "\Documentation\License.htm" 'INIPath & 
        'Else
        '
        'End If

        Help1Start = My.Application.Info.DirectoryPath & "\Documentation\Elektronische Messwert Erfassung 1_0_0_8.htm" 'INIPath & 
        Me.TopMost = False

        Debug.Print(Help1Start)

        Process.Start("IExplore.exe", Help1Start)

    End Sub

    Private Sub LizenzToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LizenzToolStripMenuItem.Click
        Dim Help1Start As String
        '
        'If State = "DE" Then
        'Help1Start = My.Application.Info.DirectoryPath & "\Documentation\License.htm" 'INIPath & 
        'Else
        '
        'End If

        Help1Start = My.Application.Info.DirectoryPath & "\Documentation\License.htm" 'INIPath & 
        Me.TopMost = False

        Debug.Print(Help1Start)

        Process.Start("IExplore.exe", Help1Start)
    End Sub
End Class