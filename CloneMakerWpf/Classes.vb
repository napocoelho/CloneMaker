Imports System.Data
Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.Xml.Serialization
Imports System.IO

Public Class Replace
    Public Property De As String
    Public Property Por As String

    Public Sub New(De As String, Por As String)
        Me.De = De
        Me.Por = Por
    End Sub

    Public Sub New()
    End Sub

End Class

Public Class Problem
    Public Property Descrição As String
    Public Property Observação As String

    Public Sub New(Descrição As String, Observação As String)
        Me.Descrição = Descrição
        Me.Observação = Observação
    End Sub

    Public Sub New()
    End Sub
End Class


Public Class Log
    Public Property Source As String
    Public Property Description As String

    Public Sub New(Source As String, Description As String)
        Me.Source = Source
        Me.Description = Description
    End Sub

    Public Sub New()
    End Sub
End Class

Public Class CloneMakerConfiguration

    Public Property Host As String
    Public Property DatabaseSource As String
    Public Property DatabaseDestination As String

    Public Property CreateDatabaseScript As Boolean
    Public Property CreateTablesScript As Boolean
    Public Property CreateDefaultConstraintsScript As Boolean
    Public Property CreateCheckConstraintsScript As Boolean
    Public Property CreateUniqueConstraintsScript As Boolean
    Public Property CreatePrimaryKeysScript As Boolean
    Public Property CreateForeignKeysScript As Boolean
    Public Property CreateStoredProceduresScript As Boolean
    Public Property CreateTriggersScript As Boolean
    Public Property CreateViewsScript As Boolean

    Public Property DoOrfanFix As Boolean
    Public Property DoComparing As Boolean
    Public Property DoStructureClone As Boolean
    Public Property DoDataClone As Boolean

    Public Property Replaces As ObservableCollection(Of Replace)
    Public Property Problems As ObservableCollection(Of Problem)
    Public Property Logs As ObservableCollection(Of Log)

    Public Property IsRunning As Boolean

    Public Sub New()
        Replaces = New ObservableCollection(Of Replace)
        Problems = New ObservableCollection(Of Problem)
        Logs = New ObservableCollection(Of Log)


        Host = String.Empty
        DatabaseSource = String.Empty
        DatabaseDestination = String.Empty

        CreateDatabaseScript = False
        CreateTablesScript = False
        CreateDefaultConstraintsScript = False
        CreateCheckConstraintsScript = False
        CreateUniqueConstraintsScript = False
        CreatePrimaryKeysScript = False
        CreateForeignKeysScript = False
        CreateStoredProceduresScript = False
        CreateTriggersScript = False
        CreateViewsScript = False

        DoOrfanFix = False
        DoComparing = False
        DoStructureClone = False
        DoDataClone = False

        Me.IsRunning = False
    End Sub


    Public Function GetSelectedCommands() As List(Of ScriptMaker.CommandType)

        Dim commandTypes As New List(Of ScriptMaker.CommandType)

        If CreateDatabaseScript Then commandTypes.Add(ScriptMaker.CommandType.Database)
        If CreateTablesScript Then commandTypes.Add(ScriptMaker.CommandType.Table)
        If CreateDefaultConstraintsScript Then commandTypes.Add(ScriptMaker.CommandType.DefaultConstraint)
        If CreateCheckConstraintsScript Then commandTypes.Add(ScriptMaker.CommandType.CheckConstraint)
        If CreateUniqueConstraintsScript Then commandTypes.Add(ScriptMaker.CommandType.UniqueConstraint)
        If CreatePrimaryKeysScript Then commandTypes.Add(ScriptMaker.CommandType.PrimaryKey)
        If CreateForeignKeysScript Then commandTypes.Add(ScriptMaker.CommandType.ForeignKey)
        If CreateStoredProceduresScript Then commandTypes.Add(ScriptMaker.CommandType.StoredProcedure)
        If CreateTriggersScript Then commandTypes.Add(ScriptMaker.CommandType.Trigger)
        If CreateViewsScript Then commandTypes.Add(ScriptMaker.CommandType.View)

        Return commandTypes

    End Function

    Public Shared Sub SaveInXml(ByRef Obj As CloneMakerConfiguration, ByVal FilePath As String)
        Dim sw As New StreamWriter(FilePath)
        Dim serializer As New XmlSerializer(GetType(CloneMakerConfiguration))
        serializer.Serialize(sw, Obj)
        sw.Close()
    End Sub


    Public Shared Function LoadFromXml(ByVal FilePath As String) As CloneMakerConfiguration
        Dim sr As New StreamReader(FilePath)
        Dim serializer As New XmlSerializer(GetType(CloneMakerConfiguration))

        Dim newObj As CloneMakerConfiguration
        newObj = serializer.Deserialize(sr)
        sr.Close()

        Return newObj
    End Function

    Public Shared Function GetInstanceFromCommandLineArgs() As CloneMakerConfiguration

        Dim NewConfig As New CloneMakerConfiguration


        'Comandos de importação:
        For Each Arg As String In Environment.GetCommandLineArgs.ToList
            If Arg.ExistsIn("/file") Then

                Dim FilePath As String = Arg.Split({"="}, StringSplitOptions.None)(1)

                NewConfig = GetInstanceFromFile(FilePath)

            End If
        Next

        'Comandos de configuração:
        For Each Arg As String In Environment.GetCommandLineArgs.ToList

            If Arg.ExistsIn("/Host") Then
                NewConfig.Host = Arg.Split({"="}, StringSplitOptions.None)(1)
            ElseIf Arg.ExistsIn("/DatabaseSource") Then
                NewConfig.DatabaseSource = Arg.Split({"="}, StringSplitOptions.None)(1)
            ElseIf Arg.ExistsIn("/DatabaseDestination") Then
                NewConfig.DatabaseDestination = Arg.Split({"="}, StringSplitOptions.None)(1)


                'Scripts:
            ElseIf Arg.ExistsIn("/CreateDatabaseScript") Then
                NewConfig.CreateDatabaseScript = CBool(Arg.Split({"="}, StringSplitOptions.None)(1))
            ElseIf Arg.ExistsIn("/CreateTablesScript") Then
                NewConfig.CreateTablesScript = CBool(Arg.Split({"="}, StringSplitOptions.None)(1))
            ElseIf Arg.ExistsIn("/CreateDefaultConstraintsScript") Then
                NewConfig.CreateDefaultConstraintsScript = CBool(Arg.Split({"="}, StringSplitOptions.None)(1))
            ElseIf Arg.ExistsIn("/CreateCheckConstraintsScript") Then
                NewConfig.CreateCheckConstraintsScript = CBool(Arg.Split({"="}, StringSplitOptions.None)(1))
            ElseIf Arg.ExistsIn("/CreateUniqueConstraintsScript") Then
                NewConfig.CreateUniqueConstraintsScript = CBool(Arg.Split({"="}, StringSplitOptions.None)(1))
            ElseIf Arg.ExistsIn("/CreatePrimaryKeysScript") Then
                NewConfig.CreatePrimaryKeysScript = CBool(Arg.Split({"="}, StringSplitOptions.None)(1))
            ElseIf Arg.ExistsIn("/CreateForeignKeysScript") Then
                NewConfig.CreateForeignKeysScript = CBool(Arg.Split({"="}, StringSplitOptions.None)(1))
            ElseIf Arg.ExistsIn("/CreateStoredProceduresScript") Then
                NewConfig.CreateStoredProceduresScript = CBool(Arg.Split({"="}, StringSplitOptions.None)(1))
            ElseIf Arg.ExistsIn("/CreateTriggersScript") Then
                NewConfig.CreateTriggersScript = CBool(Arg.Split({"="}, StringSplitOptions.None)(1))
            ElseIf Arg.ExistsIn("/CreateViewsScript") Then
                NewConfig.CreateViewsScript = CBool(Arg.Split({"="}, StringSplitOptions.None)(1))


                'Outras opções:
            ElseIf Arg.ExistsIn("/DoOrfanFix") Then
                NewConfig.DoOrfanFix = CBool(Arg.Split({"="}, StringSplitOptions.None)(1))
            ElseIf Arg.ExistsIn("/DoComparing") Then
                NewConfig.DoComparing = CBool(Arg.Split({"="}, StringSplitOptions.None)(1))
            ElseIf Arg.ExistsIn("/DoStructureClone") Then
                NewConfig.DoStructureClone = CBool(Arg.Split({"="}, StringSplitOptions.None)(1))
            ElseIf Arg.ExistsIn("/DoDataClone") Then
                NewConfig.DoDataClone = CBool(Arg.Split({"="}, StringSplitOptions.None)(1))
            End If

        Next

        Return NewConfig

    End Function

    Public Shared Function GetInstanceFromFile(ByVal FilePath As String) As CloneMakerConfiguration

        If Not System.IO.File.Exists(FilePath) Then _
                    Throw New ArgumentException("Invalid argument [{0}].".FormatTo(FilePath), "/file")

        Return CloneMakerConfiguration.LoadFromXml(FilePath)

    End Function


End Class


