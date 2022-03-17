Public Class Table
    Public Property Name As String
    Public Property Database As String
    Public Property Fields As List(Of Field)
    Public Property PrimaryKeys As List(Of Field)
    Public Property ForeignKeys As List(Of Field)

    Public Sub New()
        Me.Name = String.Empty
        Me.Database = String.Empty
        Me.Fields = New List(Of Field)
        Me.PrimaryKeys = New List(Of Field)
        Me.ForeignKeys = New List(Of Field)
    End Sub

End Class
