Imports System.Diagnostics
Imports System.Windows.Forms

Public Class Explorer1

    'Gibt an, ob der ausgewählte Knoten der Strukturansicht programmgesteuert geändert wird
    Private ChangingSelectedNode As Boolean

    Private Sub Explorer1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Benutzeroberfläche einrichten
        SetUpListViewColumns()
        LoadTree()
    End Sub

    Private Sub LoadTree()
        ' TODO: Code zum Hinzufügen von Elementen in der Strukturansicht hinzufügen

        Dim tvRoot As TreeNode
        Dim tvNode As TreeNode

        tvRoot = Me.TreeView.Nodes.Add("Root")
        tvNode = tvRoot.Nodes.Add("TreeItem1")
        tvNode = tvRoot.Nodes.Add("TreeItem2")
        tvNode = tvRoot.Nodes.Add("TreeItem3")
    End Sub

    Private Sub LoadListView()
        ' TODO: Code zum Hinzufügen von Elementen in der Listenansicht, basierend auf den ausgewählten Elementen der Strukturansicht, hinzufügen

        Dim lvItem As ListViewItem
        ListView.Items.Clear()

        lvItem = ListView.Items.Add("ListViewItem1")
        lvItem.SubItems.AddRange(New String() {"Column2", "Column3"})

        lvItem = ListView.Items.Add("ListViewItem2")
        lvItem.SubItems.AddRange(New String() {"Column2", "Column3"})

        lvItem = ListView.Items.Add("ListViewItem3")
        lvItem.SubItems.AddRange(New String() {"Column2", "Column3"})
    End Sub

    Private Sub SetUpListViewColumns()
        ' TODO: Code zum Einrichten von Spalten in der Listenansicht hinzufügen
        ListView.Columns.Add("Column1")
        ListView.Columns.Add("Column2")
        ListView.Columns.Add("Column3")
        SetView(View.Details)
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        'Anwendung beenden
        Global.System.Windows.Forms.Application.Exit()
    End Sub

    Private Sub ToolBarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolBarToolStripMenuItem.Click
        'Sichtbarkeit des Toolstrips und aktivierten Zustand des zugehörigen Menüelements umschalten
        ToolBarToolStripMenuItem.Checked = Not ToolBarToolStripMenuItem.Checked
        ToolStrip.Visible = ToolBarToolStripMenuItem.Checked
    End Sub

    Private Sub StatusBarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StatusBarToolStripMenuItem.Click
        'Sichtbarkeit des Statusstrips und aktivierten Zustand des zugehörigen Menüelements umschalten
        StatusBarToolStripMenuItem.Checked = Not StatusBarToolStripMenuItem.Checked
        StatusStrip.Visible = StatusBarToolStripMenuItem.Checked
    End Sub

    'Sichtbarkeit des Ordnerbereichs ändern
    Private Sub ToggleFoldersVisible()
        'Zuerst den aktivierten Zustand des zugehörigen Menüelements umschalten
        FoldersToolStripMenuItem.Checked = Not FoldersToolStripMenuItem.Checked

        'Symbolleistenschaltfläche "Ordner" für die Synchronisierung ändern
        FoldersToolStripButton.Checked = FoldersToolStripMenuItem.Checked

        ' Bereich mit TreeView reduzieren.
        Me.SplitContainer.Panel1Collapsed = Not FoldersToolStripMenuItem.Checked
    End Sub

    Private Sub FoldersToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FoldersToolStripMenuItem.Click
        ToggleFoldersVisible()
    End Sub

    Private Sub FoldersToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FoldersToolStripButton.Click
        ToggleFoldersVisible()
    End Sub

    Private Sub SetView(ByVal View As System.Windows.Forms.View)
        'Bestimmen, welche Menüelemente aktiviert werden sollen
        Dim MenuItemToCheck As ToolStripMenuItem = Nothing
        Select Case View
            Case View.Details
                MenuItemToCheck = DetailsToolStripMenuItem
            Case View.LargeIcon
                MenuItemToCheck = LargeIconsToolStripMenuItem
            Case View.List
                MenuItemToCheck = ListToolStripMenuItem
            Case View.SmallIcon
                MenuItemToCheck = SmallIconsToolStripMenuItem
            Case View.Tile
                MenuItemToCheck = TileToolStripMenuItem
            Case Else
                Debug.Fail("Unerwartete Ansicht")
                View = View.Details
                MenuItemToCheck = DetailsToolStripMenuItem
        End Select

        'Entsprechendes Menüelement aktivieren und Auswahl aller anderen Elemente im Menü "Ansichten" aufheben
        For Each MenuItem As ToolStripMenuItem In ListViewToolStripButton.DropDownItems
            If MenuItem Is MenuItemToCheck Then
                MenuItem.Checked = True
            Else
                MenuItem.Checked = False
            End If
        Next

        'Abschließend die angeforderte Ansicht festlegen
        ListView.View = View
    End Sub

    Private Sub ListToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListToolStripMenuItem.Click
        SetView(View.List)
    End Sub

    Private Sub DetailsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DetailsToolStripMenuItem.Click
        SetView(View.Details)
    End Sub

    Private Sub LargeIconsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LargeIconsToolStripMenuItem.Click
        SetView(View.LargeIcon)
    End Sub

    Private Sub SmallIconsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SmallIconsToolStripMenuItem.Click
        SetView(View.SmallIcon)
    End Sub

    Private Sub TileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TileToolStripMenuItem.Click
        SetView(View.Tile)
    End Sub

    Private Sub OpenToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles OpenToolStripMenuItem.Click
        Dim OpenFileDialog As New OpenFileDialog
        OpenFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        OpenFileDialog.Filter = "Textdateien (*.txt)|*.txt"
        OpenFileDialog.ShowDialog(Me)

        Dim FileName As String = OpenFileDialog.FileName
        ' TODO: Code zum Öffnen der Datei hinzufügen
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        Dim SaveFileDialog As New SaveFileDialog
        SaveFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        SaveFileDialog.Filter = "Textdateien (*.txt)|*.txt"
        SaveFileDialog.ShowDialog(Me)

        Dim FileName As String = SaveFileDialog.FileName
        ' TODO: Hier Code einfügen, um den aktuellen Inhalt des Formulars in einer Datei zu speichern.
    End Sub

    Private Sub TreeView_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView.AfterSelect
        ' TODO: Code zum Ändern des ListView-Inhalts basierend auf dem derzeit ausgewählten Knoten der Strukturansicht hinzufügen
        LoadListView()
    End Sub

End Class
