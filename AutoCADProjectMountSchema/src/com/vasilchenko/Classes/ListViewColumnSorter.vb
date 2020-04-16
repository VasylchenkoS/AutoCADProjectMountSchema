Imports System.Windows.Forms
Imports AutoCADProjectMountSchema.com.vasilchenko.Modules

Namespace com.vasilchenko.Classes

    Public Class ListViewColumnSorter
        Implements System.Collections.IComparer

        Private ReadOnly _mColumnNumber As Integer
        Private ReadOnly _mSortOrder As SortOrder

        Public Sub New(ByVal columnNumber As Integer, ByVal sortOrder As SortOrder)
            _mColumnNumber = columnNumber
            _mSortOrder = sortOrder
        End Sub
        ' Compare the items in the appropriate column 

        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
            Dim itemX As ListViewItem = DirectCast(x, ListViewItem)
            Dim itemY As ListViewItem = DirectCast(y, ListViewItem)
            ' Get the sub-item values. 
            Dim stringX As String
            Dim stringY As String

            If itemX.SubItems.Count <= _mColumnNumber Then
                stringX = ""
            Else
                stringX = itemX.SubItems(_mColumnNumber).Text
            End If

            If itemY.SubItems.Count <= _mColumnNumber Then
                stringY = ""
            Else
                stringY = itemY.SubItems(_mColumnNumber).Text
            End If

            ' Compare them. 
            If _mSortOrder = SortOrder.Ascending Then
                If IsNumeric(stringX) And IsNumeric(stringY) Then
                    Return (Val(stringX).CompareTo(Val(stringY)))
                ElseIf IsDate(stringX) And IsDate(stringY) Then
                    Return (DateTime.Parse(stringX).CompareTo(DateTime.Parse(stringY)))
                ElseIf IsNumeric(AdditionalFunctions.GetFirstNumericFromString(stringX)) And IsNumeric(AdditionalFunctions.GetFirstNumericFromString(stringY)) Then
                    Return (AdditionalFunctions.GetFirstNumericFromString(stringX).CompareTo(AdditionalFunctions.GetFirstNumericFromString(stringY)))
                Else
                    Return (String.Compare(stringX, stringY))
                End If
            Else
                If IsNumeric(stringX) And IsNumeric(stringY) Then
                    Return (Val(stringY).CompareTo(Val(stringX)))
                ElseIf IsDate(stringX) And IsDate(stringY) Then
                    Return (DateTime.Parse(stringY).CompareTo(DateTime.Parse(stringX)))
                ElseIf IsNumeric(AdditionalFunctions.GetFirstNumericFromString(stringX)) And IsNumeric(AdditionalFunctions.GetFirstNumericFromString(stringY)) Then
                    Return (AdditionalFunctions.GetFirstNumericFromString(stringY).CompareTo(AdditionalFunctions.GetFirstNumericFromString(stringX)))
                Else
                    Return (String.Compare(stringY, stringX))
                End If
            End If

        End Function

    End Class
End Namespace
