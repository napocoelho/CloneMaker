
Public Delegate Sub LogEventHandler(sender As Object, e As LogEventArgs)


Public Class LogEventArgs
    Inherits EventArgs

    Public Property Source As String
    Public Property Description As String

    Public Sub New()
    End Sub

    Public Sub New(Source As String, Description As String)
        Me.Source = Source
        Me.Description = Description
    End Sub

End Class

