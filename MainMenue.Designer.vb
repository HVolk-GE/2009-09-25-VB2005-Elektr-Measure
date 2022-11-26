<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainMenue
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Wird vom Windows Form-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Windows Form-Designer erforderlich.
    'Das Bearbeiten ist mit dem Windows Form-Designer möglich.  
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainMenue))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.Testeintrag1ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.VerschleissmessungenToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.DTVMessungToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.PhotografieVorlagenToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.InformationenToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ListeVonAktuellenTestsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ViewWearTemplatesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.KonfigToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.AnwendungKonfigurierenToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.COMPortTestToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.HilfeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.HilfeToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
        Me.StepByStepToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.KurzAnleitungToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.LizenzToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.BeendenToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.Label3 = New System.Windows.Forms.Label
        Me.File1 = New Microsoft.VisualBasic.Compatibility.VB6.FileListBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.WorkbookBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.HelpProvider1 = New System.Windows.Forms.HelpProvider
        Me.Version = New System.Windows.Forms.Label
        Me.MenuStrip1.SuspendLayout()
        CType(Me.WorkbookBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.MenuStrip1.Dock = System.Windows.Forms.DockStyle.Left
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.Testeintrag1ToolStripMenuItem, Me.InformationenToolStripMenuItem, Me.KonfigToolStripMenuItem, Me.HilfeToolStripMenuItem, Me.BeendenToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(93, 313)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        Me.MenuStrip1.TextDirection = System.Windows.Forms.ToolStripTextDirection.Vertical90
        '
        'Testeintrag1ToolStripMenuItem
        '
        Me.Testeintrag1ToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.VerschleissmessungenToolStripMenuItem, Me.DTVMessungToolStripMenuItem, Me.PhotografieVorlagenToolStripMenuItem})
        Me.Testeintrag1ToolStripMenuItem.Name = "Testeintrag1ToolStripMenuItem"
        Me.Testeintrag1ToolStripMenuItem.Size = New System.Drawing.Size(80, 17)
        Me.Testeintrag1ToolStripMenuItem.Text = "Messungen"
        Me.Testeintrag1ToolStripMenuItem.TextDirection = System.Windows.Forms.ToolStripTextDirection.Horizontal
        '
        'VerschleissmessungenToolStripMenuItem
        '
        Me.VerschleissmessungenToolStripMenuItem.Image = CType(resources.GetObject("VerschleissmessungenToolStripMenuItem.Image"), System.Drawing.Image)
        Me.VerschleissmessungenToolStripMenuItem.Name = "VerschleissmessungenToolStripMenuItem"
        Me.VerschleissmessungenToolStripMenuItem.Size = New System.Drawing.Size(186, 22)
        Me.VerschleissmessungenToolStripMenuItem.Text = "Verschleißmessung"
        '
        'DTVMessungToolStripMenuItem
        '
        Me.DTVMessungToolStripMenuItem.Image = CType(resources.GetObject("DTVMessungToolStripMenuItem.Image"), System.Drawing.Image)
        Me.DTVMessungToolStripMenuItem.Name = "DTVMessungToolStripMenuItem"
        Me.DTVMessungToolStripMenuItem.Size = New System.Drawing.Size(186, 22)
        Me.DTVMessungToolStripMenuItem.Text = "DTV Messung"
        '
        'PhotografieVorlagenToolStripMenuItem
        '
        Me.PhotografieVorlagenToolStripMenuItem.Image = CType(resources.GetObject("PhotografieVorlagenToolStripMenuItem.Image"), System.Drawing.Image)
        Me.PhotografieVorlagenToolStripMenuItem.Name = "PhotografieVorlagenToolStripMenuItem"
        Me.PhotografieVorlagenToolStripMenuItem.Size = New System.Drawing.Size(186, 22)
        Me.PhotografieVorlagenToolStripMenuItem.Text = "Photografie Vorlagen"
        '
        'InformationenToolStripMenuItem
        '
        Me.InformationenToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ListeVonAktuellenTestsToolStripMenuItem, Me.ViewWearTemplatesToolStripMenuItem})
        Me.InformationenToolStripMenuItem.Name = "InformationenToolStripMenuItem"
        Me.InformationenToolStripMenuItem.Size = New System.Drawing.Size(80, 17)
        Me.InformationenToolStripMenuItem.Text = "Informationen"
        Me.InformationenToolStripMenuItem.TextDirection = System.Windows.Forms.ToolStripTextDirection.Horizontal
        '
        'ListeVonAktuellenTestsToolStripMenuItem
        '
        Me.ListeVonAktuellenTestsToolStripMenuItem.Image = CType(resources.GetObject("ListeVonAktuellenTestsToolStripMenuItem.Image"), System.Drawing.Image)
        Me.ListeVonAktuellenTestsToolStripMenuItem.Name = "ListeVonAktuellenTestsToolStripMenuItem"
        Me.ListeVonAktuellenTestsToolStripMenuItem.Size = New System.Drawing.Size(203, 22)
        Me.ListeVonAktuellenTestsToolStripMenuItem.Text = "Liste von aktuellen Tests"
        '
        'ViewWearTemplatesToolStripMenuItem
        '
        Me.ViewWearTemplatesToolStripMenuItem.Image = CType(resources.GetObject("ViewWearTemplatesToolStripMenuItem.Image"), System.Drawing.Image)
        Me.ViewWearTemplatesToolStripMenuItem.Name = "ViewWearTemplatesToolStripMenuItem"
        Me.ViewWearTemplatesToolStripMenuItem.Size = New System.Drawing.Size(203, 22)
        Me.ViewWearTemplatesToolStripMenuItem.Text = "View Wear Templates"
        '
        'KonfigToolStripMenuItem
        '
        Me.KonfigToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AnwendungKonfigurierenToolStripMenuItem, Me.COMPortTestToolStripMenuItem})
        Me.KonfigToolStripMenuItem.Name = "KonfigToolStripMenuItem"
        Me.KonfigToolStripMenuItem.Size = New System.Drawing.Size(80, 17)
        Me.KonfigToolStripMenuItem.Text = "Konfig"
        Me.KonfigToolStripMenuItem.TextDirection = System.Windows.Forms.ToolStripTextDirection.Horizontal
        '
        'AnwendungKonfigurierenToolStripMenuItem
        '
        Me.AnwendungKonfigurierenToolStripMenuItem.Image = CType(resources.GetObject("AnwendungKonfigurierenToolStripMenuItem.Image"), System.Drawing.Image)
        Me.AnwendungKonfigurierenToolStripMenuItem.Name = "AnwendungKonfigurierenToolStripMenuItem"
        Me.AnwendungKonfigurierenToolStripMenuItem.Size = New System.Drawing.Size(208, 22)
        Me.AnwendungKonfigurierenToolStripMenuItem.Text = "Anwendung konfigurieren"
        '
        'COMPortTestToolStripMenuItem
        '
        Me.COMPortTestToolStripMenuItem.Image = CType(resources.GetObject("COMPortTestToolStripMenuItem.Image"), System.Drawing.Image)
        Me.COMPortTestToolStripMenuItem.Name = "COMPortTestToolStripMenuItem"
        Me.COMPortTestToolStripMenuItem.Size = New System.Drawing.Size(208, 22)
        Me.COMPortTestToolStripMenuItem.Text = "COM-Port Test"
        '
        'HilfeToolStripMenuItem
        '
        Me.HilfeToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.HilfeToolStripMenuItem1, Me.StepByStepToolStripMenuItem, Me.KurzAnleitungToolStripMenuItem, Me.LizenzToolStripMenuItem})
        Me.HilfeToolStripMenuItem.Name = "HilfeToolStripMenuItem"
        Me.HilfeToolStripMenuItem.Size = New System.Drawing.Size(80, 17)
        Me.HilfeToolStripMenuItem.Text = "Hilfe"
        Me.HilfeToolStripMenuItem.TextDirection = System.Windows.Forms.ToolStripTextDirection.Horizontal
        '
        'HilfeToolStripMenuItem1
        '
        Me.HilfeToolStripMenuItem1.Name = "HilfeToolStripMenuItem1"
        Me.HilfeToolStripMenuItem1.Size = New System.Drawing.Size(154, 22)
        Me.HilfeToolStripMenuItem1.Text = "Hilfe"
        '
        'StepByStepToolStripMenuItem
        '
        Me.StepByStepToolStripMenuItem.Name = "StepByStepToolStripMenuItem"
        Me.StepByStepToolStripMenuItem.Size = New System.Drawing.Size(154, 22)
        Me.StepByStepToolStripMenuItem.Text = "Step by Step"
        '
        'KurzAnleitungToolStripMenuItem
        '
        Me.KurzAnleitungToolStripMenuItem.Name = "KurzAnleitungToolStripMenuItem"
        Me.KurzAnleitungToolStripMenuItem.Size = New System.Drawing.Size(154, 22)
        Me.KurzAnleitungToolStripMenuItem.Text = "Kurz Anleitung"
        '
        'LizenzToolStripMenuItem
        '
        Me.LizenzToolStripMenuItem.Name = "LizenzToolStripMenuItem"
        Me.LizenzToolStripMenuItem.Size = New System.Drawing.Size(154, 22)
        Me.LizenzToolStripMenuItem.Text = "Lizenz"
        '
        'BeendenToolStripMenuItem
        '
        Me.BeendenToolStripMenuItem.Name = "BeendenToolStripMenuItem"
        Me.BeendenToolStripMenuItem.Size = New System.Drawing.Size(80, 17)
        Me.BeendenToolStripMenuItem.Text = "Beenden"
        Me.BeendenToolStripMenuItem.TextDirection = System.Windows.Forms.ToolStripTextDirection.Horizontal
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(96, 12)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(45, 13)
        Me.Label3.TabIndex = 17
        Me.Label3.Text = "Label3"
        '
        'File1
        '
        Me.File1.BackColor = System.Drawing.SystemColors.Window
        Me.File1.Cursor = System.Windows.Forms.Cursors.Default
        Me.File1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.File1.FormattingEnabled = True
        Me.File1.Location = New System.Drawing.Point(141, 161)
        Me.File1.Name = "File1"
        Me.File1.Pattern = "*.*"
        Me.File1.Size = New System.Drawing.Size(113, 82)
        Me.File1.TabIndex = 18
        Me.File1.Visible = False
        '
        'Button1
        '
        Me.Button1.AutoSize = True
        Me.Button1.Location = New System.Drawing.Point(141, 71)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 19
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(141, 144)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 13)
        Me.Label1.TabIndex = 20
        Me.Label1.Text = "Label1"
        '
        'WorkbookBindingSource
        '
        Me.WorkbookBindingSource.DataSource = GetType(Microsoft.Office.Interop.Excel._Workbook)
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(141, 111)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(81, 17)
        Me.CheckBox1.TabIndex = 21
        Me.CheckBox1.Text = "CheckBox1"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(239, 12)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(116, 37)
        Me.PictureBox1.TabIndex = 22
        Me.PictureBox1.TabStop = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(96, 102)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(39, 13)
        Me.Label2.TabIndex = 23
        Me.Label2.Text = "Label2"
        '
        'HelpProvider1
        '
        Me.HelpProvider1.HelpNamespace = "C:\Dokumente und Einstellungen\volkh\Eigene Dateien\Visual Studio 2005\Projects\E" & _
            "lektr-Measure\Projekt1.NET\Documentation\Elektronische Messwert Erfassung Alpha " & _
            "2.htm"
        '
        'Version
        '
        Me.Version.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Version.BackColor = System.Drawing.Color.Transparent
        Me.Version.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Version.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Version.Location = New System.Drawing.Point(96, 52)
        Me.Version.Name = "Version"
        Me.Version.Size = New System.Drawing.Size(199, 20)
        Me.Version.TabIndex = 24
        Me.Version.Text = "Version {0}.{1:00}"
        '
        'MainMenue
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(367, 313)
        Me.Controls.Add(Me.Version)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.File1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.MenuStrip1)
        Me.HelpButton = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "MainMenue"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Main Menu"
        Me.TopMost = True
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.WorkbookBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Testeintrag1ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents VerschleissmessungenToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DTVMessungToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents InformationenToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PhotografieVorlagenToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ListeVonAktuellenTestsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BeendenToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Public WithEvents File1 As Microsoft.VisualBasic.Compatibility.VB6.FileListBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents WorkbookBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents ViewWearTemplatesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents KonfigToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents COMPortTestToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HilfeToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StepByStepToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpProvider1 As System.Windows.Forms.HelpProvider
    Friend WithEvents AnwendungKonfigurierenToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents KurzAnleitungToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Version As System.Windows.Forms.Label
    Friend WithEvents HilfeToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LizenzToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
