Imports MongoDB.Bson
Imports MongoDB.Driver
Imports MongoDB.Driver.Linq

Class PagePrincipal

    Private connectionString As String = "mongodb://localhost"
    Private Property Db As MongoDatabase

    Private Sub Grid1_Loaded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles Grid1.Loaded

        Db = New MongoClient(connectionString).GetServer.GetDatabase("testes")



    End Sub


    Private Sub BtnTeste_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BtnTeste.Click

        Dim c As New Conta, c2 As New Conta


        'c.Codigo = "1"
        'c.Descrição = "Contas à receber"

        'c2.Codigo = "2"
        'c2.Descrição = "Contas à pagar"

        'Db.GetCollection("Contas").Insert(c)
        'Db.GetCollection("Contas").Insert(c2)




        'Dim Query As IQueryable(Of Conta)
        'Query = Db.GetCollection(Of Conta)("Contas").AsQueryable(Of Conta)()

        'For Each Cnt As Conta In Query
        '    Cnt.Descrição = Cnt.Descrição
        '    Stop
        'Next


        'Dim lista As New SortedSet(Of Conta)

        'lista.Add(New Conta("2.1.5.3.2", "conta 2.1.5.3.2"))

        'lista.Add(New Conta("2.1.5.1.2", "conta 2.1.5.1.2"))
        'lista.Add(New Conta("2.1.5", "conta 2.1.5"))
        'lista.Add(New Conta("2.1.4", "conta 2.1.4"))
        'lista.Add(New Conta("2.1.4.1", "conta 2.1.4.1"))

        'lista.Add(New Conta("1.6.2.1", "conta 1.6.2.1"))
        'lista.Add(New Conta("7.2", "conta 7.2"))
        'lista.Add(New Conta("2", "conta 2"))
        'lista.Add(New Conta("1", "conta 1"))

        'lista.Add(New Conta("2.1", "conta 2.1"))
        'lista.Add(New Conta("1.1", "conta 1.1"))
        'lista.Add(New Conta("3", "conta 3"))

        'lista.Add(New Conta("2.3", "conta 2.3"))
        'lista.Add(New Conta("2.2", "conta 2.2"))

        'lista.Add(New Conta("1", "conta 1"))
        
        'LstBoxContas.ItemsSource = lista


        'Stop

    End Sub
End Class
