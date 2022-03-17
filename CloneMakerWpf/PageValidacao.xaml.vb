Class PageValidacao

    Private Shadows Property Tentativas As UShort


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        Tentativas = 0
    End Sub

    Private Sub BtnSenha_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BtnSenha.Click

        Dim pwd As New System.Text.StringBuilder("")
        pwd.Append("s")
        pwd.Append("h")
        pwd.Append("i")
        pwd.Append("t")



        'Verificando senha:
        If PwdSenha.Password = pwd.ToString Then

            'Limpando memória:
            pwd.Clear()
            pwd = Nothing
            PwdSenha.Password = String.Empty

            If Date.Now <= New Date(2014, 4, 30) Then
                Me.NavigationService.Navigate(New PagePrincipal(True))
            End If
            Return
        Else

            BlkErro.Visibility = Windows.Visibility.Visible
        End If

        'Limpando memória:
        pwd.Clear()
        pwd = Nothing

        Me.Tentativas += 1

        'Fecha o sistema se tentar logar 3 vezes seguidas:
        If Tentativas < 1 Or Tentativas >= 3 Then

            MsgBox("Foram feitas 3 tentativas. O sistema será abortado.", vbCritical, "Aviso")
            Application.Current.Shutdown()
        End If

        PwdSenha.Password = String.Empty

    End Sub

    Private Sub Grid1_Loaded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles Grid1.Loaded
        PwdSenha.Focus()
    End Sub

    Private Sub Grid1_KeyDown(sender As System.Object, e As System.Windows.Input.KeyEventArgs) Handles Grid1.KeyDown
        If e.Key = Key.Enter Then
            BtnSenha_Click(BtnSenha, Nothing)
        End If
    End Sub
End Class
