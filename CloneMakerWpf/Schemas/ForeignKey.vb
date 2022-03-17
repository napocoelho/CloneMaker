Public Class ForeignKey
    Public Property Name As String
    Public Property DatabaseName As String

    Public Property TableName As String
    Public Property RelatedTableName As String

    Public Property Fields As IList(Of String)
    Public Property RelatedFields As IList(Of String)


    Public Sub New()
        Fields = New List(Of String)
        RelatedFields = New List(Of String)
    End Sub

End Class
