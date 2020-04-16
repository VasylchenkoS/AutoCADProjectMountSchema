Imports System.Windows.Forms
Imports AutoCADProjectMountSchema.com.vasilchenko.Classes
Imports AutoCADProjectMountSchema.com.vasilchenko.DatabaseConnection
Imports AutoCADProjectMountSchema.com.vasilchenko.Modules

Public Class ufLocationSelector
    Private Sub ufLocationSelector_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Dim acDrawingList As New SortedList(Of Integer, String)

        'Dim i As Short = 1
        'For Each strLocations In DatabaseDataAccessObject.GetLocations
        '    acDrawingList.Add(i, strLocations)
        '    i += 1
        'Next

        Dim acDrawingList = DatabaseDataAccessObject.GetLocations

        'lvSheets.Columns.Add("Checked", Windows.HorizontalAlignment.Left)
        lvPlaces.MultiSelect = True
        lvPlaces.FullRowSelect = True
        lvPlaces.Sorting = SortOrder.Ascending
        lvPlaces.GridLines = True
        lvPlaces.View = View.Details
        lvPlaces.AutoSize = True
        AutoSize = True

        lvPlaces.Sorting = SortOrder.Ascending
        'lvSheets.Sort()

        'lvPlaces.Columns.Add("##", 50, HorizontalAlignment.Left)
        lvPlaces.Columns.Add("Full Name", 450, HorizontalAlignment.Left)
        'lvSheets.CheckBoxes = True
        Dim arrLvItem(acDrawingList.Count - 1) As Windows.Forms.ListViewItem

        Dim i = 0
        For Each element In acDrawingList
            arrLvItem(i) = New Windows.Forms.ListViewItem(element)
            i += 1
        Next element

        'For Each pair As KeyValuePair(Of Integer, String) In acDrawingList
        '    arrLvItem(i) = New Windows.Forms.ListViewItem(pair.Key)
        '    arrLvItem(i).SubItems.Add(pair.Value)
        '    i += 1
        'Next pair

        lvPlaces.Items.AddRange(arrLvItem)

    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private _mSortingColumn As ColumnHeader

    Private Sub ListView1_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lvPlaces.ColumnClick
        ' Get the new sorting column. 
        Dim newSortingColumn As ColumnHeader = lvPlaces.Columns(e.Column)
        ' Figure out the new sorting order. 
        Dim sortOrder As SortOrder
        If _mSortingColumn Is Nothing Then
            ' New column. Sort ascending. 
            sortOrder = SortOrder.Ascending
        Else ' See if this is the same column. 
            If newSortingColumn.Equals(_mSortingColumn) Then
                ' Same column. Switch the sort order. 
                If _mSortingColumn.Text.StartsWith("> ") Then
                    sortOrder = SortOrder.Descending
                Else
                    sortOrder = SortOrder.Ascending
                End If
            Else
                ' New column. Sort ascending. 
                sortOrder = SortOrder.Ascending
            End If
            ' Remove the old sort indicator. 
            _mSortingColumn.Text = _mSortingColumn.Text.Substring(2)
        End If
        ' Display the new sort order. 
        _mSortingColumn = newSortingColumn
        If sortOrder = SortOrder.Ascending Then
            _mSortingColumn.Text = "> " & _mSortingColumn.Text
        Else
            _mSortingColumn.Text = "< " & _mSortingColumn.Text
        End If
        ' Create a comparer. 
        lvPlaces.ListViewItemSorter = New ListViewColumnSorter(e.Column, sortOrder)
        ' Sort. 
        lvPlaces.Sort()
    End Sub

    Private Sub btnApply_Click(sender As Object, e As EventArgs) Handles btnApply.Click

        If lvPlaces.SelectedItems.Count <> 1 Then
            MsgBox("Выберите одно месторасположение элементов", MsgBoxStyle.Critical)
        Else
            CreateElementList(lvPlaces.SelectedItems(0).Text)
            Me.Close()
        End If

    End Sub
End Class