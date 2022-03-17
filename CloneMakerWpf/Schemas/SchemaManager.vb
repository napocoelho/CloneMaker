Imports System.Data

Public Class SchemaManager

    Public Property Connection As ConnectionManager



    Public Sub New(ByRef Connection As ConnectionManager)
        Me.Connection = Connection
    End Sub

    Private Shadows Function GetSqlForeignKeys(ByVal DatabaseName As String) As String

        Dim StrSql As String

        StrSql = "" & _
                 "SELECT FK.CONSTRAINT_NAME       AS FkName        , " & _
                 "       KU.TABLE_NAME            AS DependentTable, " & _
                 "       KU.COLUMN_NAME           AS DependentCol  , " & _
                 "       KU.ORDINAL_POSITION      AS DependentOrder, " & _
                 "       KU2.TABLE_NAME           AS SourceTable   , " & _
                 "       KU2.COLUMN_NAME          AS SourceCol     , " & _
                 "       KU2.ORDINAL_POSITION     AS SourceOrder     " & _
                 "FROM   " & DatabaseName.Squared & ".INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS FK " & _
                 "       INNER JOIN " & DatabaseName.Squared & ".INFORMATION_SCHEMA.KEY_COLUMN_USAGE KU " & _
                 "       ON     KU.CONSTRAINT_NAME = FK.CONSTRAINT_NAME " & _
                 "       INNER JOIN " & DatabaseName.Squared & ".INFORMATION_SCHEMA.KEY_COLUMN_USAGE KU2 " & _
                 "       ON     KU2.CONSTRAINT_NAME = FK.UNIQUE_CONSTRAINT_NAME " & _
                 "       AND    KU.ORDINAL_POSITION = KU2.ORDINAL_POSITION"



        Return "SELECT * FROM ( " & StrSql & " ) AS RESULT"

    End Function


    Public Shadows Function GetForeignKey(ByVal DatabaseName As String, ByVal ForeignKeyName As String) As ForeignKey

        Dim StrSql As String, Rd As IDataReader, Retorno As ForeignKey = Nothing

        StrSql = GetSqlForeignKeys(DatabaseName) & " WHERE FkName = " & ForeignKeyName.Quote & " ORDER BY FkName, DependentOrder"


        Rd = Me.Connection.ExecuteDataReader(StrSql)

        Try

            While Rd.Read

                If Retorno Is Nothing Then
                    Retorno = New ForeignKey
                    Retorno.Name = Rd("FkName").ToString
                    Retorno.DatabaseName = DatabaseName
                    Retorno.TableName = Rd("DependentTable").ToString
                    Retorno.RelatedTableName = Rd("SourceTable").ToString
                End If

                Retorno.Fields.Add(Rd("DependentCol").ToString)
                Retorno.RelatedFields.Add(Rd("SourceCol").ToString)
            End While

        Finally
            Me.Connection.CloseDataReader()
        End Try

        Return Retorno
    End Function

    Public Shadows Function GetForeignKeys(ByVal DatabaseName As String) As IList(Of ForeignKey)

        Dim StrSql As String, Rd As IDataReader, Retorno As New List(Of ForeignKey)
        Dim UltimoFk As String = String.Empty

        StrSql = GetSqlForeignKeys(DatabaseName) & " ORDER BY FkName, DependentOrder"

        Rd = Me.Connection.ExecuteDataReader(StrSql)

        Try

            While Rd.Read

                If UltimoFk.ToUpper <> Rd("FkName").ToString.ToUpper Then

                    UltimoFk = Rd("FkName").ToString.ToUpper

                    Retorno.Add(New ForeignKey)

                    Retorno.Last.Name = Rd("FkName").ToString
                    Retorno.Last.DatabaseName = DatabaseName
                    Retorno.Last.TableName = Rd("DependentTable").ToString
                    Retorno.Last.RelatedTableName = Rd("SourceTable").ToString
                End If

                Retorno.Last.Fields.Add(Rd("DependentCol").ToString)
                Retorno.Last.RelatedFields.Add(Rd("SourceCol").ToString)
            End While

        Finally
            Me.Connection.CloseDataReader()
        End Try

        Return Retorno
    End Function

End Class
