Imports System.Data
Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Collections.Concurrent
Imports System.Threading
Imports System.Windows.Threading.DispatcherTimer


Class PageExecucao

    Private Shadows ScriptWorker As BackgroundWorker
    Private Shadows CloneWorker As BackgroundWorker

    Private Shadows QueueLog As New ConcurrentQueue(Of Log)
    Private Shadows QueueProblems As New ConcurrentQueue(Of Problem)

    Private Shadows TempoPercorrido As TimeSpan
    Private Shadows Temporizador As System.Windows.Threading.DispatcherTimer

    Private Shadows Property SavedScript As String

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


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.

        


        ProgressBar1.IsIndeterminate = False

        



        ConfigurarCronômetro()

        ConfigurarWorkers()
        
        ConfigurarHistórico()

    End Sub

    Private Sub AtivarBotoes(Ativar As Boolean)
        BtnGerarScript.IsEnabled = Ativar
        BtnClonar.IsEnabled = Ativar
    End Sub


    Private Shadows Sub ConfigurarHistórico()

        Dim DispatcherLog As New System.Windows.Threading.DispatcherTimer
        Dim DispatcherProblem As New System.Windows.Threading.DispatcherTimer

        'Iniciando verificadores de mensagens:
        AddHandler DispatcherLog.Tick, Sub()
                                           Dim LogResult As Log = Nothing

                                           While QueueLog.TryDequeue(LogResult)
                                               Config.Logs.Add(LogResult)
                                               GridLogs.ScrollIntoView(GridLogs.Items.GetItemAt(GridLogs.Items.Count - 1))
                                           End While
                                       End Sub


        AddHandler DispatcherProblem.Tick, Sub()
                                               Dim LogResult As Problem = Nothing

                                               While QueueProblems.TryDequeue(LogResult)
                                                   Config.Problems.Add(LogResult)
                                                   GridProblemas.ScrollIntoView(GridProblemas.Items.GetItemAt(GridProblemas.Items.Count - 1))
                                               End While
                                           End Sub


        DispatcherLog.Interval = New TimeSpan(3000)
        DispatcherProblem.Interval = New TimeSpan(3000)

        DispatcherLog.Start()
        DispatcherProblem.Start()

    End Sub

    Private Shadows Sub ConfigurarCronômetro()

        'Configurando temporizador (contagem do tempo de execução):
        Temporizador = New System.Windows.Threading.DispatcherTimer

        'O correto seria criar uma classe Cronômetro (ou algo parecido) para encapsular a funcionalidade
        'de contagem de tempo, mas.... ia dar mais trabalho, e deixei assim:
        AddHandler Temporizador.Tick, Sub()
                                          TempoPercorrido = TempoPercorrido.Add(TimeSpan.FromSeconds(1))
                                          LblTempoPercorrido.Content = TempoPercorrido
                                      End Sub

    End Sub

    Private Shadows Sub ConfigurarWorkers()

        ScriptWorker = New BackgroundWorker
        CloneWorker = New BackgroundWorker

        'Configurando Workers:
        ScriptWorker.WorkerSupportsCancellation = True
        CloneWorker.WorkerSupportsCancellation = True
        AddHandler ScriptWorker.DoWork, AddressOf GerarScripts
        AddHandler CloneWorker.DoWork, AddressOf ClonarBase


        AddHandler ScriptWorker.RunWorkerCompleted, Sub()

                                                        ProgressBar1.IsIndeterminate = False

                                                        Temporizador.Stop()

                                                        If Not Me.SavedScript.IsNull And Not Me.SavedScript.IsEmpty Then

                                                            MsgBox("O procedimento foi concluído." & vbNewLine & _
                                                                   "Aperte OK para que o script seja gravado no Clipboard." & vbNewLine & _
                                                                   "Para mais informações, consute o Log.", MsgBoxStyle.Information, "Procedimento finalizado")

                                                            Clipboard.Clear()
                                                            Clipboard.SetText(Me.SavedScript)
                                                        End If


                                                        AtivarBotoes(True)
                                                    End Sub

        AddHandler CloneWorker.RunWorkerCompleted, Sub()
                                                       ProgressBar1.IsIndeterminate = False

                                                       Temporizador.Stop()

                                                       MsgBox("O procedimento foi concluído." & vbNewLine & _
                                                               "Para mais informações, consute o Log.", MsgBoxStyle.Information, "Procedimento finalizado")

                                                       AtivarBotoes(True)
                                                   End Sub

    End Sub




    Private Sub GerarScripts()

        Dim script As ScriptMaker

        Dim ScriptSql As New System.Text.StringBuilder

        Try

            SavedScript = ""



            'Conectando na base:
            Conexao = ConectarServidor(Config.Host, Config.DatabaseSource)

            If Conexao.GetCurrentConnection.State = ConnectionState.Closed Then
                Conexao.GetCurrentConnection.Open()
            End If

            Conexao.GetCurrentConnection.ChangeDatabase(Config.DatabaseSource)

            'Instanciando os objetos que farão o serviço pesado (ScriptMaker e CloneMaker):
            script = New ScriptMaker(Conexao)

            script.BackgroundWorker = ScriptWorker

            AddHandler script.LogEvent, AddressOf ScriptLogHandler

            'Gerando nova base de dados:
            For Each xSql As String In script.GenerateCommandList(Config.DatabaseSource, Config.DatabaseDestination, Config.GetSelectedCommands)

                If ScriptWorker.CancellationPending Then
                    Throw New Exception("Procedimento cancelado pelo usuário")
                End If

                Try
                    Dim SqlFinal As String

                    SqlFinal = xSql

                    For Each Rep As Replace In Config.Replaces
                        SqlFinal = SqlFinal.Replace(Rep.De, Rep.Por)
                    Next

                    ScriptSql.AppendLine(SqlFinal)
                    ScriptSql.AppendLine("GO")
                    ScriptSql.AppendLine("")

                Catch ex As Exception
                    Dim Prob As New Problem
                    Prob.Descrição = ex.Message
                    Prob.Observação = xSql
                    QueueProblems.Enqueue(Prob)
                End Try

            Next



            'MsgBox("O procedimento foi finalizado e o script será gravado no Clipboard assim que pressionar OK! Para maiores informações, consulte o Log.", MsgBoxStyle.Information, "Procedimento finalizado")

            Me.SavedScript = ScriptSql.ToString

            'Clipboard.Clear()
            'Clipboard.SetText(ScriptSql.ToString)

        Catch ex As Exception

            'Dim Msg As String = ex.Message
            'MsgBox(ex.Message, vbExclamation, "Erro")

            QueueProblems.Enqueue(New Problem(ex.Message, "erro"))
            'Stop

        Finally

            If Conexao.GetCurrentConnection.State = ConnectionState.Open Then
                Conexao.GetCurrentConnection.Close()
            End If

        End Try

    End Sub


    Private Sub ClonarBase()

        Dim script As ScriptMaker
        Dim clone As CloneMaker
        Dim CommandList As List(Of String)

        Try



            'Conectando na base:
            Conexao = ConectarServidor(Config.Host, Config.DatabaseSource)

            If Conexao.GetCurrentConnection.State = ConnectionState.Closed Then
                Conexao.GetCurrentConnection.Open()
            End If

            Conexao.GetCurrentConnection.ChangeDatabase(Config.DatabaseSource)

            'Instanciando os objetos que farão o serviço pesado (ScriptMaker e CloneMaker):
            script = New ScriptMaker(Conexao)
            clone = New CloneMaker(Conexao, Config.DatabaseSource, Config.DatabaseDestination)

            script.BackgroundWorker = ScriptWorker
            clone.BackgroundWorker = CloneWorker

            AddHandler script.LogEvent, AddressOf ScriptLogHandler
            AddHandler clone.LogEvent, AddressOf CloneLogHandler



            'Concertando registros órfãos:
            If Config.DoOrfanFix Then
                clone.OrfanRecordFix()
            End If

            'Gerando nova base de dados:
            If Config.DoStructureClone Then
                CommandList = script.GenerateCommandList(Config.DatabaseSource, Config.DatabaseDestination, Config.GetSelectedCommands)

                For Each xSql As String In CommandList

                    If CloneWorker.CancellationPending Then
                        Throw New Exception("Procedimento cancelado pelo usuário")
                    End If

                    Try
                        Dim SqlFinal As String

                        SqlFinal = xSql

                        For Each Rep As Replace In Config.Replaces
                            SqlFinal = SqlFinal.Replace(Rep.De, Rep.Por)
                        Next

                        Conexao.ExecuteNonQuery(SqlFinal) '.Replace("Latin1_General_CI_AS", "Latin1_General_CI_AI").Replace("NVARCHAR", "VARCHAR").Replace("nvarchar", "varchar"))
                    Catch ex As Exception
                        Dim Prob As New Problem
                        Prob.Descrição = ex.Message
                        Prob.Observação = xSql

                        QueueProblems.Enqueue(Prob)
                    End Try

                Next
            End If


            If Config.DoDataClone Then

                'Exportando registros para nova base de dados:
                clone.Export()


                'Comparando as 2 bases:
                If Config.DoComparing Then
                    If Not clone.Comparing() Then

                        QueueProblems.Enqueue(New Problem("Uma ou mais tabelas não foram copiadas", "Verificar o log"))
                    End If
                End If

            End If

            'MsgBox("O procedimento foi finalizado! Para maiores informações, consulte o Log.", MsgBoxStyle.Information, "Procedimento finalizado")

        Catch ex As Exception

            QueueProblems.Enqueue(New Problem(ex.Message, "erro"))

            'Dim Msg As String = ex.Message
            'MsgBox(ex.Message, vbExclamation, "Erro")

        Finally

            If Conexao.GetCurrentConnection.State = ConnectionState.Open Then
                Conexao.GetCurrentConnection.Close()
            End If

        End Try


    End Sub

    Public Sub ScriptLogHandler(sender As Object, evt As LogEventArgs)

        QueueLog.Enqueue(New Log(evt.Source, evt.Description))

    End Sub

    Public Sub CloneLogHandler(sender As Object, evt As LogEventArgs)

        QueueLog.Enqueue(New Log(evt.Source, evt.Description))

    End Sub


    Private Sub BtnGerarScript_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BtnGerarScript.Click

        AtivarBotoes(False)
        ProgressBar1.IsIndeterminate = True


        'Configurando workers
        Config.Logs.Clear()
        Config.Problems.Clear()
        ConfigurarWorkers()

        'Configurando timer:
        TempoPercorrido = TimeSpan.FromSeconds(0)
        LblTempoPercorrido.Content = TempoPercorrido
        Temporizador.Interval = TimeSpan.FromSeconds(1)
        Temporizador.Start()



        'Iniciando procedimento:        
        ScriptWorker.RunWorkerAsync() 'chama GerarScripts em outra Thread

    End Sub

    Private Sub BtnClonar_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BtnClonar.Click

        AtivarBotoes(False)
        ProgressBar1.IsIndeterminate = True


        'Configurando workers
        Config.Logs.Clear()
        Config.Problems.Clear()
        ConfigurarWorkers()

        'Configurando timer:
        TempoPercorrido = TimeSpan.FromSeconds(0)
        LblTempoPercorrido.Content = TempoPercorrido
        Temporizador.Interval = TimeSpan.FromSeconds(1)
        Temporizador.Start()


        'Iniciando procedimento:        
        CloneWorker.RunWorkerAsync()  'chama ClonarBase em outra Thread
        
    End Sub

    Private Sub BtnCancelar_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BtnCancelar.Click

        If CloneWorker.IsBusy Then
            CloneWorker.CancelAsync()
        End If

        If ScriptWorker.IsBusy Then
            ScriptWorker.CancelAsync()
        End If

    End Sub


    Private Sub BtnGerarScript_IsEnabledChanged(sender As System.Object, e As System.Windows.DependencyPropertyChangedEventArgs) Handles BtnGerarScript.IsEnabledChanged
        BtnCancelar.IsEnabled = Not (BtnGerarScript.IsEnabled And BtnClonar.IsEnabled)
    End Sub

    Private Sub BtnClonar_IsEnabledChanged(sender As System.Object, e As System.Windows.DependencyPropertyChangedEventArgs) Handles BtnClonar.IsEnabledChanged
        BtnCancelar.IsEnabled = Not (BtnGerarScript.IsEnabled And BtnClonar.IsEnabled)
    End Sub

    
End Class

