Imports MongoDB.Bson
Imports MongoDB.Driver
Imports MongoDB.Driver.Linq
Imports System.Collections.Generic

Public Class Conta
    Implements IComparable(Of Conta)

    Private Shadows _codigo As String

    Public Property Id As ObjectId
    Public Property Descrição As String

    Public Property Código As String
        Get
            Return _codigo
        End Get
        Set(value As String)

            Dim Classes() As String = value.Split({"."}, StringSplitOptions.None)

            'Validando conteúdo do código:
            For Each Classe As String In Classes
                For Each Digit As Char In Classe.ToArray
                    If Not Char.IsNumber(Digit) Then
                        Throw New ArgumentException("O código inserido é inválido!", "Conta.Código")
                    End If
                Next
            Next

            Me._codigo = value
        End Set
    End Property


    Public Property PertenceA As Conta
    Public Property Possui As SortedSet(Of Conta)

    Public Sub New()
    End Sub

    Public Sub New(Código As String, Descrição As String)
        Me.Código = Código
        Me.Descrição = Descrição
    End Sub

    Public Function CompareTo(ByVal other As Finanças.Conta) As Integer Implements IComparable(Of Finanças.Conta).CompareTo
        Return Me.Código.CompareTo(other.Código)
    End Function

    
    Public Overrides Function ToString() As String
        Return Me.Código
    End Function

End Class

Public Class PlanoDeContas

    Public Property Id As ObjectId
    Public Property Descrição As String

    Public Property Contas As SortedSet(Of Conta)


    Public Sub New()
    End Sub

End Class
