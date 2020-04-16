Namespace com.vasilchenko.Classes
    Public Class ConnectionClass
        Public Property ConnectionPin As PinClass
        Public Property ConnectionTo As List(Of PinClass)
        Public Property BlockTagNumber As String
        Public Function ConnnectionToAsString() As String
            Return String.Join("; ", ConnectionTo.ConvertAll(Function(x) x.PinDescription))
        End Function
    End Class
End Namespace

