Option Strict Off
Option Explicit On
Friend Class Form1
    Inherits System.Windows.Forms.Form
    Public CntBefore, CntAfter, CntDS2Before, CntDS2After As Integer

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: Diese Codezeile lädt Daten in die Tabelle "MessungenDataSet1.VorlagenFormat". Sie können sie bei Bedarf verschieben oder entfernen.

        Me.VorlagenFormatTableAdapter.Fill(Me.MessungenDataSet1.VorlagenFormat)
        'INSERT INTO VorlagenFormat (Templatename, [Measurement Pads Points], _ 
        '[Measurement Disc Points], [Measurement Pads Times], [Measurement Disc Times]) _
        'VALUES (?, ?, ?, ?, ?)

        CntBefore = 0
        CntDS2Before = 0

        con.ConnectionString = strCon

        sqladapterVor = New OleDb.OleDbDataAdapter("Select * from VorlagenFormat", con)

        con.Open()

        sqladapterVor.Fill(DS2)

        CntDS2Before = DS2.Tables(0).Rows.Count

        CntBefore = Me.VorlagenFormatDataGridView.RowCount

        CntBefore = CntBefore - 1

        If State = "DE" Then
            Me.Button1.Text = "Schliessen"
            Me.Button2.Text = "Erstellen"
            Me.Label1.Text = "Templatename"
            Me.Label2.Text = "Measurement" & Chr(10) & "Pads Points"
            Me.Label3.Text = "Measurement" & Chr(10) & "Disc Points"
            Me.Label4.Text = "Measurement" & Chr(10) & "Pads Times"
            Me.Label5.Text = "Measurement" & Chr(10) & "Disc Times"
        Else
            Me.Button1.Text = "Close"
            Me.Button2.Text = "Create"
            Me.Label1.Text = "Templatename"
            Me.Label2.Text = "Measurement" & Chr(10) & "Pads Points"
            Me.Label3.Text = "Measurement" & Chr(10) & "Disc Points"
            Me.Label4.Text = "Measurement" & Chr(10) & "Pads Times"
            Me.Label5.Text = "Measurement" & Chr(10) & "Disc Times"
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        con.Close()
        Me.Hide()
        ApplConfig.Show()
    End Sub

    '#########################################################################################

    Private Sub VorlagenFormatBindingNavigatorSaveItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VorlagenFormatBindingNavigatorSaveItem.Click
       
        'UpdateData()

        Me.Validate()
        Me.VorlagenFormatBindingSource.EndEdit()
        'Me.VorlagenFormatTableAdapter.Insert 

        'AddNewOrders()
        UpdateData()


        Me.MessungenDataSet1.AcceptChanges()
        'MessungenDataSet1.AcceptChanges()

    End Sub

    Private Sub UpdateData()

        Me.Validate()
        Me.VorlagenFormatBindingSource.EndEdit()

        Using updateTransaction As New Transactions.TransactionScope

            DeleteOrders()
            AddNewOrders()

            updateTransaction.Complete()

            Me.MessungenDataSet1.AcceptChanges()

        End Using
    End Sub

    Private Sub DeleteOrders()
        Dim deletedOrders As MessungenDataSet.VorlagenFormatDataTable ' TemplateDataTable
        deletedOrders = CType(MessungenDataSet1.VorlagenFormat.GetChanges(Data.DataRowState.Deleted), _
             MessungenDataSet.VorlagenFormatDataTable)

        If Not IsNothing(deletedOrders) Then
            Try

                ' sqladapterVor.Update(deletedOrders)
                Me.VorlagenFormatTableAdapter.Update(deletedOrders)

            Catch ex As Exception
                MessageBox.Show("Delete Failed")
            End Try
        End If
    End Sub

    Private Sub AddNewOrders()

        Dim newOrders As MessungenDataSet.VorlagenFormatDataTable
        newOrders = CType(MessungenDataSet1.VorlagenFormat.GetChanges(Data.DataRowState.Added), _
             MessungenDataSet.VorlagenFormatDataTable)

        If Not IsNothing(newOrders) Then
            Try
                'sqladapterVor.Update(newOrders)
                VorlagenFormatTableAdapter.Update(newOrders)
                'VorlagenFormatTableAdapter.Insert(newOrders)

            Catch ex As Exception
                MessageBox.Show("AddNew Failed" & vbCrLf & ex.Message)
            End Try
        End If
    End Sub

    '#########################################################################################

    Private Sub InitializeDataGridView()
        Try
            ' Set up the DataGridView.
            With Me.VorlagenFormatDataGridView
                ' Automatically generate the DataGridView columns.
                .AutoGenerateColumns = True

                ' Set up the data source.
                VorlagenFormatBindingSource.DataSource = GetData("Select * From VorlagenFormat")
                .DataSource = VorlagenFormatBindingSource
                ' Automatically resize the visible rows.
                .AutoSizeRowsMode = _
                    DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders

                ' Set the DataGridView control's border.
                .BorderStyle = BorderStyle.Fixed3D

                ' Put the cells in edit mode when user enters them.
                .EditMode = DataGridViewEditMode.EditOnEnter
            End With
        Catch ex As SqlClient.SqlException
            MessageBox.Show("To run this sample replace " _
                & "connection.ConnectionString with a valid connection string" _
                & "  to a Northwind database accessible to your system.", _
                "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            System.Threading.Thread.CurrentThread.Abort()
        End Try
    End Sub

    Private Shared Function GetData(ByVal sqlCommand As String) _
        As DataTable

        Dim connectionString As String = _
            "Integrated Security=SSPI;Persist Security Info=False;" _
            & "Initial Catalog=Messungen;Data Source=localhost"

        Dim northwindConnection As SqlClient.SqlConnection = _
            New SqlClient.SqlConnection(connectionString)

        Dim command As New SqlClient.SqlCommand(sqlCommand, northwindConnection)
        Dim adapter As SqlClient.SqlDataAdapter = New SqlClient.SqlDataAdapter()
        adapter.SelectCommand = command

        Dim table As New DataTable
        table.Locale = System.Globalization.CultureInfo.InvariantCulture
        adapter.Fill(table)

        Return table

    End Function
    '###########################################################################


    '#########################################################################################

    Sub TempKeineFunktion()
        Dim tmpTemplaename As String
        Dim MPP, MDP, MPT, MDT, IDNr, i As Integer
        Dim strArrayBefore(0, 4) As String
        Dim strArrayAfter(0, 4) As String


        If CntBefore = CntAfter Then

            CntBefore = CntBefore - 1
            CntAfter = CntAfter - 1
            ReDim strArrayBefore(CntBefore, 4)
            ReDim strArrayAfter(CntBefore, 4)

            For i = 0 To CntBefore
                strArrayBefore(i, 0) = Me.VorlagenFormatDataGridView.Rows(i).Cells(0).Value
                strArrayBefore(i, 1) = Me.VorlagenFormatDataGridView.Rows(i).Cells(1).Value
                strArrayBefore(i, 2) = Me.VorlagenFormatDataGridView.Rows(i).Cells(2).Value
                strArrayBefore(i, 3) = Me.VorlagenFormatDataGridView.Rows(i).Cells(3).Value
                strArrayBefore(i, 4) = Me.VorlagenFormatDataGridView.Rows(i).Cells(4).Value
            Next

            For i = 0 To CntBefore
                strArrayAfter(i, 0) = DS2.Tables.Item(0).Rows(i).Item(0).ToString
                strArrayAfter(i, 1) = DS2.Tables.Item(0).Rows(i).Item(1).ToString
                strArrayAfter(i, 2) = DS2.Tables.Item(0).Rows(i).Item(2).ToString
                strArrayAfter(i, 3) = DS2.Tables.Item(0).Rows(i).Item(3).ToString
                strArrayAfter(i, 4) = DS2.Tables.Item(0).Rows(i).Item(4).ToString
            Next

            For i = 0 To CntBefore
                If strArrayBefore(i, 0) <> strArrayAfter(i, 0) Then
                    tmpTemplaename = strArrayAfter(i, 0)
                    IDNr = i
                End If
                If strArrayBefore(i, 1) <> strArrayAfter(i, 1) Then
                    MPP = strArrayAfter(i, 1)
                    IDNr = i
                End If
                If strArrayBefore(i, 2) <> strArrayAfter(i, 2) Then
                    MDP = strArrayAfter(i, 2)
                    IDNr = i
                End If
                If strArrayBefore(i, 3) <> strArrayAfter(i, 3) Then
                    MPT = strArrayAfter(i, 3)
                    IDNr = i
                End If
                If strArrayBefore(i, 4) <> strArrayAfter(i, 4) Then
                    MDT = strArrayAfter(i, 4)
                    IDNr = i
                End If
                If IDNr > 0 Then Exit For
            Next
        End If
    End Sub
    '#########################################################################################

End Class