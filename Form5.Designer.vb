<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form5
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form5))
        Me.MessungenDataSet = New MeasureAndWeigh.MessungenDataSet
        Me.DuplTestNumbersAndTimesBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.DuplTestNumbersAndTimesTableAdapter = New MeasureAndWeigh.MessungenDataSetTableAdapters.DuplTestNumbersAndTimesTableAdapter
        Me.DuplTestNumbersAndTimesBindingNavigator = New System.Windows.Forms.BindingNavigator(Me.components)
        Me.BindingNavigatorMoveFirstItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorMovePreviousItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorSeparator = New System.Windows.Forms.ToolStripSeparator
        Me.BindingNavigatorPositionItem = New System.Windows.Forms.ToolStripTextBox
        Me.BindingNavigatorCountItem = New System.Windows.Forms.ToolStripLabel
        Me.BindingNavigatorSeparator1 = New System.Windows.Forms.ToolStripSeparator
        Me.BindingNavigatorMoveNextItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorMoveLastItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorSeparator2 = New System.Windows.Forms.ToolStripSeparator
        Me.BindingNavigatorAddNewItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorDeleteItem = New System.Windows.Forms.ToolStripButton
        Me.DuplTestNumbersAndTimesBindingNavigatorSaveItem = New System.Windows.Forms.ToolStripButton
        Me.DuplTestNumbersAndTimesDataGridView = New System.Windows.Forms.DataGridView
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.MessungenDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DuplTestNumbersAndTimesBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DuplTestNumbersAndTimesBindingNavigator, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.DuplTestNumbersAndTimesBindingNavigator.SuspendLayout()
        CType(Me.DuplTestNumbersAndTimesDataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MessungenDataSet
        '
        Me.MessungenDataSet.DataSetName = "MessungenDataSet"
        Me.MessungenDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'DuplTestNumbersAndTimesBindingSource
        '
        Me.DuplTestNumbersAndTimesBindingSource.DataMember = "DuplTestNumbersAndTimes"
        Me.DuplTestNumbersAndTimesBindingSource.DataSource = Me.MessungenDataSet
        '
        'DuplTestNumbersAndTimesTableAdapter
        '
        Me.DuplTestNumbersAndTimesTableAdapter.ClearBeforeFill = True
        '
        'DuplTestNumbersAndTimesBindingNavigator
        '
        Me.DuplTestNumbersAndTimesBindingNavigator.AddNewItem = Me.BindingNavigatorAddNewItem
        Me.DuplTestNumbersAndTimesBindingNavigator.BindingSource = Me.DuplTestNumbersAndTimesBindingSource
        Me.DuplTestNumbersAndTimesBindingNavigator.CountItem = Me.BindingNavigatorCountItem
        Me.DuplTestNumbersAndTimesBindingNavigator.DeleteItem = Me.BindingNavigatorDeleteItem
        Me.DuplTestNumbersAndTimesBindingNavigator.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BindingNavigatorMoveFirstItem, Me.BindingNavigatorMovePreviousItem, Me.BindingNavigatorSeparator, Me.BindingNavigatorPositionItem, Me.BindingNavigatorCountItem, Me.BindingNavigatorSeparator1, Me.BindingNavigatorMoveNextItem, Me.BindingNavigatorMoveLastItem, Me.BindingNavigatorSeparator2, Me.BindingNavigatorAddNewItem, Me.BindingNavigatorDeleteItem, Me.DuplTestNumbersAndTimesBindingNavigatorSaveItem})
        Me.DuplTestNumbersAndTimesBindingNavigator.Location = New System.Drawing.Point(0, 0)
        Me.DuplTestNumbersAndTimesBindingNavigator.MoveFirstItem = Me.BindingNavigatorMoveFirstItem
        Me.DuplTestNumbersAndTimesBindingNavigator.MoveLastItem = Me.BindingNavigatorMoveLastItem
        Me.DuplTestNumbersAndTimesBindingNavigator.MoveNextItem = Me.BindingNavigatorMoveNextItem
        Me.DuplTestNumbersAndTimesBindingNavigator.MovePreviousItem = Me.BindingNavigatorMovePreviousItem
        Me.DuplTestNumbersAndTimesBindingNavigator.Name = "DuplTestNumbersAndTimesBindingNavigator"
        Me.DuplTestNumbersAndTimesBindingNavigator.PositionItem = Me.BindingNavigatorPositionItem
        Me.DuplTestNumbersAndTimesBindingNavigator.Size = New System.Drawing.Size(664, 25)
        Me.DuplTestNumbersAndTimesBindingNavigator.TabIndex = 0
        Me.DuplTestNumbersAndTimesBindingNavigator.Text = "BindingNavigator1"
        '
        'BindingNavigatorMoveFirstItem
        '
        Me.BindingNavigatorMoveFirstItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveFirstItem.Image = CType(resources.GetObject("BindingNavigatorMoveFirstItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveFirstItem.Name = "BindingNavigatorMoveFirstItem"
        Me.BindingNavigatorMoveFirstItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveFirstItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMoveFirstItem.Text = "Erste verschieben"
        '
        'BindingNavigatorMovePreviousItem
        '
        Me.BindingNavigatorMovePreviousItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMovePreviousItem.Image = CType(resources.GetObject("BindingNavigatorMovePreviousItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMovePreviousItem.Name = "BindingNavigatorMovePreviousItem"
        Me.BindingNavigatorMovePreviousItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMovePreviousItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMovePreviousItem.Text = "Vorherige verschieben"
        '
        'BindingNavigatorSeparator
        '
        Me.BindingNavigatorSeparator.Name = "BindingNavigatorSeparator"
        Me.BindingNavigatorSeparator.Size = New System.Drawing.Size(6, 25)
        '
        'BindingNavigatorPositionItem
        '
        Me.BindingNavigatorPositionItem.AccessibleName = "Position"
        Me.BindingNavigatorPositionItem.AutoSize = False
        Me.BindingNavigatorPositionItem.Name = "BindingNavigatorPositionItem"
        Me.BindingNavigatorPositionItem.Size = New System.Drawing.Size(50, 21)
        Me.BindingNavigatorPositionItem.Text = "0"
        Me.BindingNavigatorPositionItem.ToolTipText = "Aktuelle Position"
        '
        'BindingNavigatorCountItem
        '
        Me.BindingNavigatorCountItem.Name = "BindingNavigatorCountItem"
        Me.BindingNavigatorCountItem.Size = New System.Drawing.Size(44, 13)
        Me.BindingNavigatorCountItem.Text = "von {0}"
        Me.BindingNavigatorCountItem.ToolTipText = "Die Gesamtanzahl der Elemente."
        '
        'BindingNavigatorSeparator1
        '
        Me.BindingNavigatorSeparator1.Name = "BindingNavigatorSeparator"
        Me.BindingNavigatorSeparator1.Size = New System.Drawing.Size(6, 6)
        '
        'BindingNavigatorMoveNextItem
        '
        Me.BindingNavigatorMoveNextItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveNextItem.Image = CType(resources.GetObject("BindingNavigatorMoveNextItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveNextItem.Name = "BindingNavigatorMoveNextItem"
        Me.BindingNavigatorMoveNextItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveNextItem.Size = New System.Drawing.Size(23, 20)
        Me.BindingNavigatorMoveNextItem.Text = "Nächste verschieben"
        '
        'BindingNavigatorMoveLastItem
        '
        Me.BindingNavigatorMoveLastItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveLastItem.Image = CType(resources.GetObject("BindingNavigatorMoveLastItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveLastItem.Name = "BindingNavigatorMoveLastItem"
        Me.BindingNavigatorMoveLastItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveLastItem.Size = New System.Drawing.Size(23, 20)
        Me.BindingNavigatorMoveLastItem.Text = "Letzte verschieben"
        '
        'BindingNavigatorSeparator2
        '
        Me.BindingNavigatorSeparator2.Name = "BindingNavigatorSeparator"
        Me.BindingNavigatorSeparator2.Size = New System.Drawing.Size(6, 6)
        '
        'BindingNavigatorAddNewItem
        '
        Me.BindingNavigatorAddNewItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorAddNewItem.Image = CType(resources.GetObject("BindingNavigatorAddNewItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorAddNewItem.Name = "BindingNavigatorAddNewItem"
        Me.BindingNavigatorAddNewItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorAddNewItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorAddNewItem.Text = "Neu hinzufügen"
        '
        'BindingNavigatorDeleteItem
        '
        Me.BindingNavigatorDeleteItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorDeleteItem.Image = CType(resources.GetObject("BindingNavigatorDeleteItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorDeleteItem.Name = "BindingNavigatorDeleteItem"
        Me.BindingNavigatorDeleteItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorDeleteItem.Size = New System.Drawing.Size(23, 20)
        Me.BindingNavigatorDeleteItem.Text = "Löschen"
        '
        'DuplTestNumbersAndTimesBindingNavigatorSaveItem
        '
        Me.DuplTestNumbersAndTimesBindingNavigatorSaveItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.DuplTestNumbersAndTimesBindingNavigatorSaveItem.Enabled = False
        Me.DuplTestNumbersAndTimesBindingNavigatorSaveItem.Image = CType(resources.GetObject("DuplTestNumbersAndTimesBindingNavigatorSaveItem.Image"), System.Drawing.Image)
        Me.DuplTestNumbersAndTimesBindingNavigatorSaveItem.Name = "DuplTestNumbersAndTimesBindingNavigatorSaveItem"
        Me.DuplTestNumbersAndTimesBindingNavigatorSaveItem.Size = New System.Drawing.Size(23, 20)
        Me.DuplTestNumbersAndTimesBindingNavigatorSaveItem.Text = "Daten speichern"
        '
        'DuplTestNumbersAndTimesDataGridView
        '
        Me.DuplTestNumbersAndTimesDataGridView.AutoGenerateColumns = False
        Me.DuplTestNumbersAndTimesDataGridView.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn1, Me.DataGridViewTextBoxColumn2, Me.DataGridViewTextBoxColumn3})
        Me.DuplTestNumbersAndTimesDataGridView.DataSource = Me.DuplTestNumbersAndTimesBindingSource
        Me.DuplTestNumbersAndTimesDataGridView.Location = New System.Drawing.Point(12, 28)
        Me.DuplTestNumbersAndTimesDataGridView.Name = "DuplTestNumbersAndTimesDataGridView"
        Me.DuplTestNumbersAndTimesDataGridView.Size = New System.Drawing.Size(345, 220)
        Me.DuplTestNumbersAndTimesDataGridView.TabIndex = 1
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.DataPropertyName = "TestnumberFeld"
        Me.DataGridViewTextBoxColumn1.HeaderText = "TestnumberFeld"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.DataPropertyName = "AnzahlVonDuplikaten"
        Me.DataGridViewTextBoxColumn2.HeaderText = "AnzahlVonDuplikaten"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        '
        'DataGridViewTextBoxColumn3
        '
        Me.DataGridViewTextBoxColumn3.DataPropertyName = "Measuretimepoint"
        Me.DataGridViewTextBoxColumn3.HeaderText = "Measuretimepoint"
        Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
        '
        'Form5
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(664, 327)
        Me.Controls.Add(Me.DuplTestNumbersAndTimesDataGridView)
        Me.Controls.Add(Me.DuplTestNumbersAndTimesBindingNavigator)
        Me.Name = "Form5"
        Me.Text = "Form5"
        CType(Me.MessungenDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DuplTestNumbersAndTimesBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DuplTestNumbersAndTimesBindingNavigator, System.ComponentModel.ISupportInitialize).EndInit()
        Me.DuplTestNumbersAndTimesBindingNavigator.ResumeLayout(False)
        Me.DuplTestNumbersAndTimesBindingNavigator.PerformLayout()
        CType(Me.DuplTestNumbersAndTimesDataGridView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MessungenDataSet As MeasureAndWeigh.MessungenDataSet
    Friend WithEvents DuplTestNumbersAndTimesBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents DuplTestNumbersAndTimesTableAdapter As MeasureAndWeigh.MessungenDataSetTableAdapters.DuplTestNumbersAndTimesTableAdapter
    Friend WithEvents DuplTestNumbersAndTimesBindingNavigator As System.Windows.Forms.BindingNavigator
    Friend WithEvents BindingNavigatorAddNewItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorCountItem As System.Windows.Forms.ToolStripLabel
    Friend WithEvents BindingNavigatorDeleteItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorMoveFirstItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorMovePreviousItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorSeparator As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents BindingNavigatorPositionItem As System.Windows.Forms.ToolStripTextBox
    Friend WithEvents BindingNavigatorSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents BindingNavigatorMoveNextItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorMoveLastItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents DuplTestNumbersAndTimesBindingNavigatorSaveItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents DuplTestNumbersAndTimesDataGridView As System.Windows.Forms.DataGridView
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn3 As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
