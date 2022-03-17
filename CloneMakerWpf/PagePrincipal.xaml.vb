Imports System.Data
Imports System.Collections.Generic
Imports System.Collections.ObjectModel

Class PagePrincipal

    Public Property RemoveBackEntry As Boolean

    Public Property Conexao As ConnectionManager
        Get
            Return MainWindow.Conexao
        End Get

        Set(value As ConnectionManager)
            MainWindow.Conexao = value
        End Set
    End Property

    Public Property Config As CloneMakerConfiguration
        Get
            Return MainWindow.Config
        End Get
        Set(value As CloneMakerConfiguration)
            MainWindow.Config = value
        End Set
    End Property

    Private Sub Page_Loaded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded

        'Reseta navegação apenas quando a origem for [PageValidacao]:
        If Me.RemoveBackEntry Then

            If Me.NavigationService.CanGoBack Then
                Me.NavigationService.RemoveBackEntry()
            End If

            Me.RemoveBackEntry = False
        End If

    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub


    Public Sub New(ByVal RemoveBackEntry As Boolean)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        Me.RemoveBackEntry = RemoveBackEntry

    End Sub

    Private Sub BtnListarBancos_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BtnListarBancos.Click

        Dim TblBancos As DataTable

        Try

            If Not Conexao Is Nothing Then
                If Conexao.GetCurrentConnection.State = ConnectionState.Open Then
                    Conexao.GetCurrentConnection.Close()
                End If
            End If


            Conexao = ConectarServidor(TxtServidor.Text, "master")

            If Conexao.GetCurrentConnection.State = ConnectionState.Closed Then
                Conexao.GetCurrentConnection.Open()
                Conexao.GetCurrentConnection.ChangeDatabase("master")
            End If


            TblBancos = Conexao.ExecuteDataTable("SELECT [name] FROM [sys].[databases] ORDER BY [name]")

            'Preenchendo combo:
            'CbxBancos.DataContext = TblBancos.DefaultView
            CbxBancos.ItemsSource = (From r As DataRow In TblBancos.Rows
                                     Select r.Item("name")).ToList

            Conexao.GetCurrentConnection.Close()

        Catch ex As Exception
            MsgBox(ex.Message, vbExclamation, "Erro")
        End Try

    End Sub

    Private Sub CbxBancos_SelectionChanged(sender As System.Object, e As System.Windows.Controls.SelectionChangedEventArgs) Handles CbxBancos.SelectionChanged
        Dim Combo As ComboBox = sender

        If Not Combo.SelectedItem Is Nothing Then
            TxtNovoBanco.Text = (Combo.SelectedItem.ToString & "_New").Trim
            Config.DatabaseDestination = (Combo.SelectedItem.ToString & "_New").Trim
        End If

        'Validando botão:
        btnProximo.IsEnabled = (Combo.Items.Count > 0 And Not Combo.SelectedItem Is Nothing)

    End Sub

    Private Sub BtnInserir_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BtnInserir.Click

        For Each Item As Replace In Config.Replaces.ToList
            If Not Item.IsNull Then
                If Item.De.IsEmpty And Item.Por.IsEmpty Then
                    Config.Replaces.Remove(Item)
                End If
            End If
        Next

        Config.Replaces.Add(New Replace)

    End Sub

    Private Sub BtnExcluir_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BtnExcluir.Click

        Dim selected As Replace = GridSubstituicoes.SelectedItem

        Config.Replaces.Remove(selected)

    End Sub

    Private Sub BtnSalvar_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BtnSalvar.Click

        Dim Dialog As New Microsoft.Win32.SaveFileDialog
        Dim ConteudoArquivo As New List(Of String)
        Dim Writer As System.IO.StreamWriter

        Dialog.Title = "Salvar arquivo..."
        Dialog.ShowDialog()

        Try

            Config.Replaces.ToList.ForEach(Sub(Item) ConteudoArquivo.Add("{0}={1}".FormatTo(Item.De, Item.Por)))

            Writer = System.IO.File.CreateText(Dialog.FileName)
            Writer.Write(ConteudoArquivo.JoinWith(vbNewLine))
            Writer.Close()

        Catch ex As Exception
            MsgBox(ex.Message, vbExclamation, "Erro")
        End Try

    End Sub

    Private Sub BtnAbrir_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BtnAbrir.Click

        Dim Dialog As New Microsoft.Win32.OpenFileDialog
        Dim ConteudoArquivo As New List(Of String)
        Dim Reader As System.IO.StreamReader

        Dialog.Title = "Abrir arquivo..."
        Dialog.ShowDialog()

        Try

            Reader = System.IO.File.OpenText(Dialog.FileName)

            ConteudoArquivo = Reader.ReadToEnd.Split({vbNewLine}, StringSplitOptions.RemoveEmptyEntries).ToList

            Config.Replaces.Clear()

            'Recuperando os parâmetros e adicionando à [Substituicoes]:
            For Each Param As String In ConteudoArquivo
                Dim Args() As String = Param.Split({"="}, StringSplitOptions.RemoveEmptyEntries)
                Dim Subs As New Replace

                Subs.De = Args(0)
                Subs.Por = Args(1)

                Config.Replaces.Add(Subs)
            Next

        Catch ex As Exception
            MsgBox(ex.Message, vbExclamation, "Erro")
        End Try

    End Sub


    


    'Private Sub BtnClonar_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BtnClonar.Click
    'BtnClonar.IsEnabled = False
    'BtnGerarScript.IsEnabled = False

    'Config.Problems.Clear()
    'ClonarBase()

    'BtnClonar.IsEnabled = True
    'BtnGerarScript.IsEnabled = True

    'End Sub

    'Private Sub BtnGerarScript_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BtnGerarScript.Click
    'BtnClonar.IsEnabled = False
    'BtnGerarScript.IsEnabled = False

    'Config.Problems.Clear()
    'GerarScripts()

    'BtnClonar.IsEnabled = True
    'BtnGerarScript.IsEnabled = True
    'End Sub


    Private Sub DoEvents()

        Dim f As New System.Windows.Threading.DispatcherFrame

        System.Windows.Threading.Dispatcher.CurrentDispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Background,
                                                                          Sub(arg As Object)
                                                                              Dim fr As System.Windows.Threading.DispatcherFrame = arg
                                                                              fr.Continue = False
                                                                          End Sub, f)
        System.Windows.Threading.Dispatcher.PushFrame(f)

    End Sub

    Private Sub btnProximo_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles btnProximo.Click
        Me.NavigationService.Navigate(New Uri("PageExecucao.xaml", UriKind.RelativeOrAbsolute))
    End Sub

    
    Private Sub BtnLoadConfig_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BtnLoadConfig.Click

        Dim Dialog As New Microsoft.Win32.OpenFileDialog
        Dim newConf As CloneMakerConfiguration

        Dialog.Title = "Abrir arquivo..."
        Dialog.FileName = "Config.xml"
        Dialog.Filter = "Xml files (*.xml)|*.*"
        Dialog.ShowDialog()

        Try
            If System.IO.File.Exists(Dialog.FileName) Then

                newConf = CloneMakerConfiguration.LoadFromXml(Dialog.FileName)

                Me.DataContext = newConf

            End If
        Catch ex As Exception
            MsgBox(ex.Message, vbExclamation, "Erro")
        End Try

    End Sub

    Private Sub BtnSaveConfig_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BtnSaveConfig.Click

        Dim Dialog As New Microsoft.Win32.SaveFileDialog

        Dialog.Title = "Salvar arquivo..."
        Dialog.FileName = "Config.xml"
        Dialog.Filter = "Xml files (*.xml)|*.*"
        Dialog.ShowDialog()

        Try
            If System.IO.File.Exists(Dialog.FileName) Then _
                CloneMakerConfiguration.SaveInXml(Me.Config, Dialog.FileName)
        Catch ex As Exception
            MsgBox(ex.Message, vbExclamation, "Erro")
        End Try

    End Sub
End Class
