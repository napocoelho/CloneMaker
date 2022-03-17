'Imports CloneMakerWpf.Utils
Imports System.Data
Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.IO


Class MainWindow

    Public Shared Property Conexao As ConnectionManager
    Public Shared Property Config As CloneMakerConfiguration
    'Public Shared Property PermissaoLiberada As Boolean

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub Window_Loaded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded

        Config = New CloneMakerConfiguration

        Me.DataContext = Config

        Config.Host = "SERVIDOR"
        Config.CreateDatabaseScript = True
        Config.CreateTablesScript = True
        Config.CreateDefaultConstraintsScript = True
        Config.CreateCheckConstraintsScript = True
        Config.CreateUniqueConstraintsScript = True
        Config.CreatePrimaryKeysScript = True
        Config.CreateForeignKeysScript = True
        Config.CreateStoredProceduresScript = True
        Config.CreateTriggersScript = True
        Config.CreateViewsScript = True

        Config.DoOrfanFix = True
        Config.DoComparing = True
        Config.DoStructureClone = True
        Config.DoDataClone = True



        'Frame = Me.frmPrincipal
        'Frame.NavigationService = 




        'Try
        '    For Each Drv As DriveInfo In DriveInfo.GetDrives

        '        Try

        '            Dim FileName As String = Drv.Name & "\dolly.txt"
        '            'Dim FileValue As String = String.Empty

        '            If File.Exists(FileName) Then

        '                'Using sr As StreamReader = New StreamReader(FileName)
        '                '    FileValue = sr.ReadToEnd
        '                'End Using

        '                'Dim DateValue As Date = Date.Parse(FileValue)


        '                PermissaoLiberada = True
        '                Exit For
        '            End If

        '        Catch ex As Exception
        '        End Try

        '    Next
        'Catch ex As Exception
        'End Try


        'If Not PermissaoLiberada Then
        '    MsgBox("Sem permissão para acesso!", vbCritical, "Aviso")
        '    Application.Current.Shutdown()
        'End If






        If Environment.GetCommandLineArgs.Length > 1 Then
            Me.Navigate(New PageExecucaoProgramada(True))
        End If


    End Sub

End Class




