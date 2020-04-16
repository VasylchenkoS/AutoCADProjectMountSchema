Namespace com.vasilchenko.Classes
    Public Class PinClass
        Public Property TagName As String
        Public Property Pin As String
        Public Property CableName As String
        Public Property Wireno As String
        Public ReadOnly Property PinDescription As String
            Get
                Return String.Format("{0}:{1}", TagName, Pin)
            End Get
        End Property
    End Class
End Namespace

