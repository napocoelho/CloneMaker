Class PageExecucaoProgramada

    Private Shadows Property RemoveBackEntry As Boolean


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Public Sub New(RemoveBackEntry As Boolean)

        ' This call is required by the designer.
        InitializeComponent()

        Me.RemoveBackEntry = RemoveBackEntry
    End Sub

    Private Sub Grid1_Loaded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles Grid1.Loaded

        'Reseta navegação apenas quando a origem for [PageValidacao]:
        If Me.RemoveBackEntry Then

            If Me.NavigationService.CanGoBack Then
                Me.NavigationService.RemoveBackEntry()
            End If

            Me.RemoveBackEntry = False
        End If




        Environment.GetCommandLineArgs.ToList.Exists(Function(x) x = x)




    End Sub

    

End Class


