Imports System.Data



Public Class CloneMaker

    Public Event LogEvent As LogEventHandler


    'Private ConnSource As ADODB.Connection, ConnDestination As ADODB.Connection
    Private ConnHlp As ConnectionManager
    Private DbSourceName As String
    Private DbDestinationName As String

    Private TableProcessingMap As Dictionary(Of String, Boolean)
    Private TableDeletingMap As Dictionary(Of String, Boolean)

    Private MapErrors As Dictionary(Of String, Boolean)

    Private BooPause As Boolean, BooStop As Boolean, BooProcessing As Boolean

    Public Property BackgroundWorker As System.ComponentModel.BackgroundWorker


    Public Sub New(ByRef ConnectionArg As ConnectionManager, _
                    ByVal DatabaseSourceNameArg As String, _
                    ByVal DatabaseDestinationNameArg As String)

        Me.BackgroundWorker = Nothing

        TableProcessingMap = New Dictionary(Of String, Boolean)
        MapErrors = New Dictionary(Of String, Boolean)

        Me.ConnHlp = ConnectionArg

        Me.DbSourceName = DatabaseSourceNameArg
        Me.DbDestinationName = DatabaseDestinationNameArg

    End Sub


    Public ReadOnly Property Connection As ConnectionManager
        Get
            Return Me.ConnHlp
        End Get
    End Property

    Public ReadOnly Property DatabaseSource As String
        Get
            Return Me.DbSourceName
        End Get
    End Property

    Public ReadOnly Property DatabaseDestination As String
        Get
            Return Me.DbDestinationName
        End Get
    End Property







    'Public Property Let TryToPauseProcess(ByVal PauseArg As Boolean)
    '   BooPause = PauseArg
    'End Property

    'Public Property Let TryToStopProcess(ByVal StopArg As Boolean)
    '   BooStop = StopArg
    'End Property


    Private ReadOnly Property IsComplete As Boolean
        Get
            For Each xItem As KeyValuePair(Of String, Boolean) In Me.TableProcessingMap
                If Not xItem.Value Then
                    Return False
                End If
            Next

            Return True
        End Get
    End Property



    '---:CAMPOS ANTIGOS:------------------------------------------------------------------------------------------------------------------------------------------------------------------


    'Exporta o banco inteiro ou as tabelas especificadas em [TableNameList]:
    Public Sub Export(Optional ByRef TableNameList As IList(Of String) = Nothing, _
                      Optional ByRef TableNameExceptionList As IList(Of String) = Nothing)

        'Dim StrSql As String
        'Dim LngSavedIndex As Long
        'Dim ConnSrcHlp As New ConnectionHelper 'Source connection
        'Dim ConnDstHlp As New ConnectionHelper 'Destination connection

        'ConnSrcHlp.Connection = Me.ConnectionSource
        'ConnDstHlp.Connection = Me.ConnectionDestination





        'Se já iniciou o processamento, continuará do ponto onde parou:
        If Not Me.BooProcessing Then



            Call ExportPreparing(TableNameList, TableNameExceptionList)

            Me.BooProcessing = True
        End If





        '-PROCESSANDO EXPORTAÇÃO---------------------------------------------------------------------------------------------------------------------------------
        For Each NomeTabela As String In TableProcessingMap.Keys.ToList

            VerificarSeUsuarioCancelou

            Try

                RaiseEvent LogEvent(Me, New LogEventArgs("Export", "Exporting table [{0}]".FormatTo(NomeTabela)))

                CopiaHierarquica(NomeTabela)
            Catch ex As Exception
                RaiseEvent LogEvent(Me, New LogEventArgs("Export problem", ex.Message))
            End Try
        Next


        'TableProcessingMap.Keys.ToList.ForEach(Sub(Item)
        '                                           'Importando:
        '                                           CopiaHierarquica(Item)
        '                                       End Sub)

    End Sub


    Private Sub ExportPreparing(Optional ByRef TableNameList As List(Of String) = Nothing, _
                                Optional ByRef TableNameExceptionList As List(Of String) = Nothing)

        Dim StrSql As String


        RaiseEvent LogEvent(Me, New LogEventArgs("ExportPreparing", "Preparing to export"))

        '-VERIFICANDO CONTEÚDO A SER PROCESSADO------------------------------------------------------------------------------------------------------------------

        'Preparando lista de tabelas a serem exportadas:
        If TableNameList Is Nothing Then

            TableNameList = New List(Of String)


            StrSql = " SELECT name FROM [{0}].[sys].[objects] " & _
                     " WHERE type = 'U' " & _
                     " ORDER BY name"

            For Each xRow As DataRow In Me.Connection.ExecuteDataTable(StrSql.FormatTo(Me.DatabaseSource.Trim)).Rows
                TableNameList.Add(xRow("name"))
            Next

        End If



        'Eliminando exceções:
        If Not TableNameExceptionList Is Nothing Then
            For Each xItem As String In TableNameExceptionList
                TableNameList.Remove(xItem)
            Next
        End If





        '-VERIFICANDO SE TABELAS EXISTEM NO DESTINO--------------------------------------------------------------------------------------------------------------
        For Each xItem As String In TableNameList.ToList
            If Not Me.TableExists(Me.DatabaseDestination, xItem) Then
                TableNameList.Remove(xItem)
            End If
        Next






        '-INICIALIZANDO NOME DAS TABELAS PARA PROCESSAMENTO------------------------------------------------------------------------------------------------------
        TableProcessingMap = New Dictionary(Of String, Boolean)
        TableDeletingMap = New Dictionary(Of String, Boolean)

        For Each xTabela As String In TableNameList.ToList
            TableProcessingMap.Add(xTabela, False)
            TableDeletingMap.Add(xTabela, False)
        Next




        '-LIMPANDO TABELAS DESTINO-------------------------------------------------------------------------------------------------------------------------------



        For Each xTabela As String In TableDeletingMap.Keys.ToList

            Try
                DeleteHierarquico(xTabela)
                'Call TruncateTable(Me.DatabaseDestination, Item)
                'Call DeleteTable(Me.DatabaseDestination, Item)
            Catch ex As Exception
                Stop
            End Try

        Next




    End Sub


    Private Function DatabaseExists(ByVal DatabaseName As String) As Boolean
        Dim StrSql As String


        StrSql = " SELECT COUNT(name) AS Total " & _
                 " FROM   [master].[sys].[databases] " & _
                 " WHERE  xtype = 'U' " & _
                 "        AND name = {1}"

        Return (CInt(Me.Connection.ExecuteScalar(StrSql.FormatTo(DatabaseName.Quote))) > 0)


    End Function


    Private Function TableExists(ByVal DatabaseName As String, _
                                  ByVal StrNomeTabela As String) As Boolean

        Dim StrSql As String



        StrSql = " SELECT COUNT(name) AS Total " & _
                 " FROM   [{0}].[sys].[objects] " & _
                 " WHERE  type = 'U' " & _
                 "        AND name = {1}"



        Return (CInt(Me.Connection.ExecuteScalar(StrSql.FormatTo(DatabaseName, StrNomeTabela.Quote))) > 0)



    End Function

    Private Function TruncateTable(ByVal DatabaseName As String, _
                                     ByVal TableName As String) As Boolean

        Try
            Call Connection.ExecuteNonQuery("TRUNCATE TABLE [{0}]..[{1}]".FormatTo(DatabaseName, TableName))
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    Private Function DeleteTable(ByVal DatabaseName As String, _
                                  ByVal TableName As String) As Boolean

        Try
            Call Connection.ExecuteNonQuery("DELETE FROM [{0}]..[{1}]".FormatTo(DatabaseName, TableName))
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function


    'Private Function CopiarTabelaSemIdentity(ByVal StrNomeTabela As String, _
    '                                          ByVal StrColunas As String) As Boolean

    '    Dim StrSql As String ', StrMsg As String


    '    StrSql = " INSERT INTO {0}..{1} ( {2} )" & _
    '         "    SELECT {2} FROM {3}..{1}"

    '    StrSql = StrSql.FormatTo(Me.DatabaseDestination.Squared, _
    '                                StrNomeTabela.Squared, _
    '                                StrColunas, _
    '                                Me.DatabaseSource.Squared)

    '    'Clipboard.Clear
    '    'Clipboard.SetText (StrSql)

    '    Call Me.Connection.ExecuteNonQuery(StrSql)
    '    Return True

    'End Function







    'Private Function CopiarTabelaComIdentity(ByVal StrNomeTabela As String, _
    '                                          ByVal StrColunas As String) As Boolean

    '    Dim StrSql As String ', StrMsg As String



    '    StrSql = " SET IDENTITY_INSERT {0}..{1} ON" & _
    '         "    INSERT INTO {0}..{1} ( {2} )" & _
    '         "       SELECT {2} FROM {3}..{1}" & _
    '         " SET IDENTITY_INSERT {0}..{1} OFF"

    '    StrSql = StrSql.FormatTo(Me.DatabaseDestination.Squared, _
    '                                       StrNomeTabela.Squared, _
    '                                       StrColunas, _
    '                                       Me.DatabaseSource.Squared)

    '    'Clipboard.Clear
    '    'Clipboard.SetText StrSql

    '    Connection.BeginTransaction()
    '    Call Connection.ExecuteNonQuery(StrSql)
    '    Connection.CommitTransaction()

    '    Return True

    'End Function




    Private Sub CopiarTabela(ByVal StrNomeTabela As String)

        CopiarTabela(Me.DatabaseSource, StrNomeTabela, Me.DatabaseDestination, StrNomeTabela)

    End Sub

    Private Sub CopiarTabela(ByVal FromDatabase As String, ByVal FromTable As String, _
                                  ByVal ToDatabase As String, ByVal ToTable As String)

        Dim ColsFrom As List(Of String), ColsTo As List(Of String), StrSql As String

        'Organizando colunas:
        ColsFrom = GetTableFields(FromDatabase, FromTable, True, True)
        ColsTo = ColsFrom.ToList
        'ColsTo = GetTableFields(ToDatabase, ToTable, True, True)


        'Preparando para fazer a cópia:
        If TableHasIdentity(ToDatabase, ToTable) Then

            StrSql = " SET IDENTITY_INSERT {0}..{1} ON" & _
                     "    INSERT INTO {0}..{1} ( {2} )" & _
                     "       SELECT {5} FROM {3}..{4} with(nolock)" & _
                     " SET IDENTITY_INSERT {0}..{1} OFF"

            StrSql = StrSql.FormatTo(ToDatabase.Squared, _
                                        ToTable.Squared, _
                                        ColsTo.JoinWith(", "), _
                                        FromDatabase.Squared, _
                                        FromTable.Squared, _
                                        ColsFrom.JoinWith(", "))




        Else

            StrSql = " INSERT INTO {0}..{1} ( {2} )" & _
                     "    SELECT {5} FROM {3}..{4} with(nolock)"

            StrSql = StrSql.FormatTo(ToDatabase.Squared, _
                                        ToTable.Squared, _
                                        ColsTo.JoinWith(", "), _
                                        FromDatabase.Squared, _
                                        FromTable.Squared, _
                                        ColsFrom.JoinWith(", "))

        End If


        'Copiando informações:
        Connection.BeginTransaction()
        Connection.ExecuteNonQuery(StrSql)
        'Call Connection.ExecuteNonQuery(StrSql)
        Connection.CommitTransaction()

    End Sub



    Private Function TableHasIdentity(ByVal DatabaseNameArg As String, ByVal TableNameArg As String) As Boolean

        Dim StrSql As String

        StrSql = String.Empty

        StrSql = StrSql & _
                 "SELECT SCHEMA_NAME(UID)           AS SCHEMA_NAME , " & _
                 "       sysobjects.ID              AS TABLE_ID    , " & _
                 "       sysobjects.name            AS TABLE_NAME  , " & _
                 "       COLUMNS_1.name             AS COLUMN_NAME , " & _
                 "       COLUMNS_1.COLID            AS COLUMN_ORDER, " & _
                 "       TYPE_NAME(COLUMNS_1.xtype) AS TYPE_NAME   , " & _
                 "       Collation                                 , " & _
                 "       CollationId                               , " & _
                 "       Prec AS PRECISION                         , " & _
                 "       COLUMNS_1.Scale                           , " & _
                 "       IsNullable                                , " & _
                 "       IsComputed                                , "

        StrSql = StrSql & _
                 "       Is_RowGuidCol                  AS IsRowGuidCol             , " & _
                 "       Is_Identity                    AS IsIdentity               , " & _
                 "       Is_FileStream                  AS IsFileStream             , " & _
                 "       Is_Replicated                  AS IsReplicated             , " & _
                 "       Is_Non_Sql_Subscribed          AS IsNonSqlSubscribed       , " & _
                 "       Is_Ansi_Padded                 AS IsAnsiPadded             , " & _
                 "       Is_Merge_Published             AS IsMergePublished         , " & _
                 "       Default_Object_Id              AS DF_CONSTRAINT_ID         , " & _
                 "       OBJECT_NAME(Default_Object_Id) AS DF_CONSTRAINT_NAME       , " & _
                 "       Rule_Object_Id                                             , " & _
                 "       Is_Sparse                                                  , " & _
                 "       Is_Column_Set                                              , " & _
                 "       Is_Xml_Document                                            , " & _
                 "       Xml_Collection_Id "

        StrSql = StrSql & _
                 "FROM   " & DatabaseNameArg.Squared & "..[sysobjects] AS sysobjects " & _
                 "       INNER JOIN " & DatabaseNameArg.Squared & "..[syscolumns] AS COLUMNS_1 " & _
                 "       ON     COLUMNS_1.Id = sysobjects.id " & _
                 "       INNER JOIN " & DatabaseNameArg.Squared & ".[sys].[columns] AS COLUMNS_2 " & _
                 "       ON     COLUMNS_2.object_id = COLUMNS_1.id " & _
                 "       AND    COLUMNS_2.column_id = COLUMNS_1.colid " & _
                 "WHERE  sysobjects.XTYPE      = 'U'" & _
                 "       AND Is_Identity       = 1 " & _
                 "       /*AND SCHEMA_NAME(UID)  = */ " & _
                 "       AND sysobjects.name   = " & TableNameArg.Quote

        'Clipboard.Clear
        'Clipboard.SetText StrSql

        StrSql = "SELECT COUNT(Table_Name) AS Total FROM ( " & StrSql & " ) AS RESULT"

        Try
            Return (CInt(Me.Connection.ExecuteScalar(StrSql)) > 0)
        Catch ex As Exception
            Return False
        End Try

    End Function

    Private Function GetPrimaryKeyFields(ByVal DatabaseName As String, ByVal TableName As String) As List(Of String)

        Dim StrSql As String, RsOrigem As New DataTable
        Dim ListaCampos As New List(Of String)


        StrSql = ""
        StrSql = StrSql & "SELECT Table_Schema             AS [Schema]  , "
        StrSql = StrSql & "       Constraint_Catalog       AS DbName    , "
        StrSql = StrSql & "       Table_Name               AS TableName , "
        StrSql = StrSql & "       constraint_name          AS PkName    , "
        StrSql = StrSql & "       Column_Name              AS ColumnName, "
        StrSql = StrSql & "       Ordinal_Position         AS ColumnOrder "
        StrSql = StrSql & "FROM   {0}.[INFORMATION_SCHEMA].[KEY_COLUMN_USAGE] "
        StrSql = StrSql & "WHERE  OBJECTPROPERTY(OBJECT_ID(constraint_name), 'IsPrimaryKey') = 1 "
        StrSql = StrSql & "         AND Table_Name = {1} "
        'StrSql = StrSql & "ORDER BY Constraint_Catalog, Table_Name, Ordinal_Position"

        StrSql = StrSql.FormatTo(DatabaseName.Squared, TableName.Quote)

        StrSql = "SELECT ColumnName FROM (" & StrSql & ") AS VIEW_CONSTRAINTS_PK ORDER BY DbName, TableName, ColumnOrder"


        RsOrigem = Me.Connection.ExecuteDataTable(StrSql)

        'ObterListaDeColunas = String.Empty

        For Each xRow As DataRow In RsOrigem.Rows
            ListaCampos.Add(xRow("ColumnName").ToString.Squared)
        Next

        Return ListaCampos

    End Function


    Private Function GetTableFields(ByVal DatabaseName As String, ByVal TableName As String, _
                                    Optional ByVal TimeStampTypeExclude As Boolean = False, _
                                    Optional ByVal ComputedFieldsExclude As Boolean = False) As List(Of String)

        Dim StrSql As String, RsOrigem As New DataTable
        Dim ListaCampos As New List(Of String)
        Dim ListaCondicoes As New List(Of String)


        ListaCondicoes.Add("TABELAS.name = " & TableName.Quote)

        If TimeStampTypeExclude Then
            ListaCondicoes.Add("COLUNAS.xtype NOT IN ( SELECT xtype FROM " & DatabaseName.Squared & "..[systypes] WHERE name IN ( 'timestamp' ) )")
        End If

        If ComputedFieldsExclude Then
            ListaCondicoes.Add("IsComputed = 0")
        End If


        StrSql = " SELECT   DISTINCT COLUNAS.name " & _
                 " FROM     " & DatabaseName.Squared & "..[sysobjects] AS TABELAS " & _
                 "          INNER JOIN " & DatabaseName.Squared & "..[syscolumns] AS COLUNAS " & _
                 "          ON COLUNAS.id = TABELAS.id " & _
                 " WHERE    " & ListaCondicoes.JoinWith(" AND ")

        'LEFT JOIN " & SquareBracket(DatabaseName) & "..[systypes] AS TIPOS ON COLUNAS.xtype = TIPOS.xtype

        RsOrigem = Me.Connection.ExecuteDataTable(StrSql)

        'ObterListaDeColunas = String.Empty

        For Each xRow As DataRow In RsOrigem.Rows
            ListaCampos.Add(xRow("name").ToString.Squared)
        Next

        'ObterListaDeColunas = ListaCampos.JoinWith(", ")

        Return ListaCampos

    End Function

    ' Cópia Hierárquica
    Private Function CopiaHierarquica(ByVal StrNomeTabela As String, _
                                      Optional ByVal BooRaiz As Boolean = True, _
                                      Optional ByRef MapHierarquia As Dictionary(Of String, Boolean) = Nothing) As Boolean

        Dim StrSql As String ', RsDependencias As DataTable
        Dim BooTodasPendenciasRetornaramOk As Boolean
        Dim TblDependencies As DataTable
        'Dim BooReturnValue As Boolean


        VerificarSeUsuarioCancelou

        'Se a [StrNomeTabela] atual já estiver copiada, apenas retorna TRUE:
        If TableProcessingMap.Item(StrNomeTabela) Then

            Return True

        Else  'Caso contrário, busca por tabelas [pendentes]:

            'O IF abaixo é importante, pois evita que aconteça loops infinitos:
            If BooRaiz Then

                MapHierarquia = New Dictionary(Of String, Boolean)

            Else

                'Evita entrar em loop:
                If MapHierarquia.ContainsKey(StrNomeTabela) Then

                    Return False
                End If

            End If

            'Adicionando hierarquia:
            MapHierarquia.Add(StrNomeTabela, False)





            'Pesquisa dependências da tabela [StrNomeTabela]:


            'StrSql = " SELECT DISTINCT " & _
            '"        OBJECT_NAME(Parent_Object_ID) AS Dependente, " & _
            '"        OBJECT_NAME(Referenced_Object_ID) AS Pendente " & _
            '" FROM   {0}.[sys].[foreign_keys] " & _
            '" WHERE  OBJECT_NAME(Parent_Object_ID) = {1}"



            StrSql = "" & _
                     " SELECT DISTINCT /*** o distinct é essencial ***/ " & _
                     "        /*KU.TABLE_NAME        AS Dependente,*/   " & _
                     "        KU2.TABLE_NAME       AS Pendente          " & _
                     "        " & _
                     "        " & _
                     "        " & _
                     "        "

            StrSql = StrSql & _
                     " FROM   {0}.INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS FK " & _
                     "        INNER JOIN {0}.INFORMATION_SCHEMA.KEY_COLUMN_USAGE KU " & _
                     "        ON     KU.CONSTRAINT_NAME = FK.CONSTRAINT_NAME " & _
                     "        INNER JOIN {0}.INFORMATION_SCHEMA.KEY_COLUMN_USAGE KU2 " & _
                     "        ON     KU2.CONSTRAINT_NAME = FK.UNIQUE_CONSTRAINT_NAME " & _
                     "        AND    KU.ORDINAL_POSITION = KU2.ORDINAL_POSITION " & _
                     " WHERE  KU.TABLE_NAME              = {1}"


            StrSql = StrSql.FormatTo(Me.DatabaseDestination.Squared, StrNomeTabela.Quote)




            'Clipboard.Clear
            'Clipboard.SetText StrSql

            'Set RsDependencias = Connection.Execute(StrSql, SquareBracket(Me.DatabaseDestination), Quote(StrNomeTabela))



            TblDependencies = Me.Connection.ExecuteDataTable(StrSql)


            BooTodasPendenciasRetornaramOk = True


            For Each xRow As DataRow In TblDependencies.Rows
                'Todas devem estar OK:
                If Not StrNomeTabela.Equals(xRow("Pendente")) Then
                    BooTodasPendenciasRetornaramOk = (BooTodasPendenciasRetornaramOk And CopiaHierarquica(xRow("Pendente"), False, MapHierarquia))
                Else
                    BooTodasPendenciasRetornaramOk = (BooTodasPendenciasRetornaramOk And True)
                End If
            Next

        End If





        'Verifica se já importou a tabela:
        If BooTodasPendenciasRetornaramOk And Not TableProcessingMap.Item(StrNomeTabela) Then

            CopiarTabela(StrNomeTabela)
            TableProcessingMap.Item(StrNomeTabela) = True
            Return True

        End If


        Return False


    End Function


    ' Cópia Hierárquica
    Private Function DeleteHierarquico(ByVal StrNomeTabela As String, _
                                      Optional ByVal BooRaiz As Boolean = True, _
                                      Optional ByRef MapHierarquia As Dictionary(Of String, Boolean) = Nothing) As Boolean

        Dim StrSql As String    ', 'RsDependencias As DataTable
        Dim BooTodasPendenciasRetornaramOk As Boolean
        Dim TblDependencies As DataTable
        Dim BooReturnValue As Boolean


        VerificarSeUsuarioCancelou


        'Se a [StrNomeTabela] atual já estiver copiada, apenas retorna TRUE:
        If CBool(TableDeletingMap.Item(StrNomeTabela)) Then

            Return True

        Else  'Caso contrário, busca por tabelas [pendentes]:

            'O IF abaixo é importante, pois evita que aconteça loops infinitos:
            If BooRaiz Then

                MapHierarquia = New Dictionary(Of String, Boolean)

            Else

                'Evita entrar em loop:
                If MapHierarquia.ContainsKey(StrNomeTabela) Then

                    Return False
                End If
            End If

            'Adicionando hierarquia:
            MapHierarquia.Add(StrNomeTabela, False)





            'Pesquisa dependências da tabela [StrNomeTabela]:


            'StrSql = " SELECT DISTINCT " & _
            '"        OBJECT_NAME(Parent_Object_ID) AS Dependente, " & _
            '"        OBJECT_NAME(Referenced_Object_ID) AS Pendente " & _
            '" FROM   {0}.[sys].[foreign_keys] " & _
            '" WHERE  OBJECT_NAME(Parent_Object_ID) = {1}"



            StrSql = "" & _
                     " SELECT DISTINCT /*** o distinct é essencial ***/ " & _
                     "        KU.TABLE_NAME        AS Dependente        " & _
                     "        /*KU2.TABLE_NAME       AS Pendente*/      " & _
                     "        " & _
                     "        " & _
                     "        " & _
                     "        "

            StrSql = StrSql & _
                     " FROM   {0}.INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS FK " & _
                     "        INNER JOIN {0}.INFORMATION_SCHEMA.KEY_COLUMN_USAGE KU " & _
                     "        ON     KU.CONSTRAINT_NAME = FK.CONSTRAINT_NAME " & _
                     "        INNER JOIN {0}.INFORMATION_SCHEMA.KEY_COLUMN_USAGE KU2 " & _
                     "        ON     KU2.CONSTRAINT_NAME = FK.UNIQUE_CONSTRAINT_NAME " & _
                     "        AND    KU.ORDINAL_POSITION = KU2.ORDINAL_POSITION " & _
                     " WHERE  KU2.TABLE_NAME              = {1}"


            StrSql = StrSql.FormatTo(Me.DatabaseDestination.Squared, StrNomeTabela.Quote)


            'Clipboard.Clear
            'Clipboard.SetText StrSql

            TblDependencies = Me.Connection.ExecuteDataTable(StrSql)


            BooTodasPendenciasRetornaramOk = True

            For Each xRow As DataRow In TblDependencies.Rows
                'Todas devem estar OK:
                If Not StrNomeTabela.Equals(xRow("Dependente")) Then
                    BooTodasPendenciasRetornaramOk = (BooTodasPendenciasRetornaramOk And DeleteHierarquico(xRow("Dependente"), False, MapHierarquia))
                Else
                    BooTodasPendenciasRetornaramOk = (BooTodasPendenciasRetornaramOk And True)
                End If
            Next

        End If




        If BooTodasPendenciasRetornaramOk Then

            'Verifica se já importou a tabela:
            If Not CBool(TableDeletingMap.Item(StrNomeTabela)) Then
                BooReturnValue = DeleteTable(Me.DatabaseDestination, StrNomeTabela)
                TableDeletingMap.Item(StrNomeTabela) = BooReturnValue
            End If



            Return BooReturnValue

        End If

        Return False

    End Function

    Public Function Comparing() As Boolean

        Dim StrSql As String, StrSelect1 As String, StrSelect2 As String

        Dim BooReturn As Boolean = True

        For Each xItem As KeyValuePair(Of String, Boolean) In TableProcessingMap

            VerificarSeUsuarioCancelou

            StrSql = "SELECT 1 WHERE ( {0} ) = ( {1} )"
            StrSelect1 = "SELECT COUNT(*) FROM " & Me.DatabaseSource.Squared & ".." & xItem.Key.Squared
            StrSelect2 = "SELECT COUNT(*) FROM " & Me.DatabaseDestination.Squared & ".." & xItem.Key.Squared

            If Me.Connection.ExecuteNonQuery(StrSql.FormatTo(StrSelect1, StrSelect2)) = 0 Then
                RaiseEvent LogEvent(Me, New LogEventArgs("Comparing", "Comparing tables {0} and {1}".FormatTo(Me.DatabaseSource.Squared & "." & xItem.Key.Squared, Me.DatabaseDestination.Squared & "." & xItem.Key.Squared)))
                BooReturn = False
            End If
        Next

        Return BooReturn
    End Function



    Public Function OrfanRecordFix() As List(Of String)

        Dim Schema As SchemaManager, FkList As IList(Of ForeignKey)
        Dim OrfanTableList As New List(Of String)
        Dim HasOrfanRecords As Boolean = False

        'OrfanTableList = New List(Of ForeignKey)
        Schema = New SchemaManager(Me.Connection)
        FkList = Schema.GetForeignKeys(Me.DatabaseSource)


        RaiseEvent LogEvent(Me, New LogEventArgs("OrfanRecordFix", "Searching for orfan records"))


        For Each FkItem As ForeignKey In FkList

            Dim StrSql As String ', SelectFieldList As New List(Of String)
            Dim FirstConditionList As New List(Of String)
            Dim SecondConditionList As New List(Of String)
            Dim SelectFieldList As New List(Of String)


            VerificarSeUsuarioCancelou


            'Concatenando partes da instruçõe sql:
            For Idx As Integer = 0 To FkItem.Fields.Count - 1
                Dim StrInner As String
                StrInner = "[DependentTbl].{0} = [RelatedTbl].{1}".FormatTo(
                    FkItem.Fields(Idx).Squared,
                    FkItem.RelatedFields(Idx).Squared
                    )

                FirstConditionList.Add(StrInner)
                SecondConditionList.Add(" AND " & FkItem.Fields(Idx).Squared & " IS NOT NULL")
                'SelectFieldList.Add(FkItem.Fields(Idx).Squared)
            Next




            'Instrução SQL que verifica se há algum registro órfão:
            StrSql = " SELECT Count({3}) AS Total " & _
                     " FROM   {0}..{1} AS [DependentTbl] " & _
                             " WHERE  NOT EXISTS ( " & _
                             "                     SELECT  1 " & _
                             "                     FROM    {0}..{2} AS [RelatedTbl] " & _
                             "                     WHERE   {4} " & _
                             "                   )" & _
                             "        {5}"

            StrSql = StrSql.FormatTo(FkItem.DatabaseName.Squared, _
                                    FkItem.TableName.Squared, _
                                    FkItem.RelatedTableName.Squared, _
                                    FkItem.Fields.First.Squared, _
                                    FirstConditionList.JoinWith(" AND "), _
                                    SecondConditionList.JoinWith(" "))


            'Clipboard.Clear()
            'Clipboard.SetText(StrSql)


            'Caso existam registros órfãos, faz um backup da informação:
            If CInt(Me.Connection.ExecuteScalar(StrSql)) > 0 Then

                Dim StrMainTable As String, StrNewTable As String
                Dim Script As ScriptMaker
                Dim StrSqlList As List(Of String)


                HasOrfanRecords = True

                StrSqlList = New List(Of String)


                Script = New ScriptMaker(Me.Connection)





                StrMainTable = FkItem.TableName.Squared
                StrNewTable = (FkItem.TableName & "_ORFAN_" & Date.Now.ToShortDateString.Replace("/", "_")).Squared


                RaiseEvent LogEvent(Me, New LogEventArgs("OrfanRecordFix", "Orfan records found in [{0}]. Moving inconsistent records from [{0}] to [{1}]".FormatTo(StrMainTable, StrNewTable)))

                'StrSqlList = New List(Of String)



                'Gerando comando que criará a nova tabela de backups dos registros órfãos:
                StrSql = Script.MakeTables(FkItem.TableName).First.Replace(StrMainTable, StrNewTable)
                StrSqlList.Add(StrSql)
                'Me.Connection.ExecuteNonQuery(StrSql)
                'Me.CopiarTabela(Me.DatabaseSource, StrMainTable, Me.DatabaseSource, StrNewTable)





                'Gerando comando que copiará registros órfãos para a tabela de bakup:
                SelectFieldList = GetTableFields(FkItem.DatabaseName, FkItem.TableName)     '--> obtendo todos os campos da tabela



                StrSql = " INSERT INTO {0}..{1} ( {2} )                         ".FormatTo(FkItem.DatabaseName.Squared, StrNewTable, SelectFieldList.JoinWith(", ")) & _
                         "      SELECT {0}                                      ".FormatTo(SelectFieldList.JoinWith(", ")) & _
                         "      FROM   {0}..{1} AS [DependentTbl]               ".FormatTo(FkItem.DatabaseName.Squared, FkItem.TableName.Squared) & _
                         "      WHERE  NOT EXISTS (                             " & _
                         "                          SELECT  1                   " & _
                         "                           FROM   {0}..{1} AS [RelatedTbl]    ".FormatTo(FkItem.DatabaseName.Squared, FkItem.RelatedTableName.Squared) & _
                         "                          WHERE   {0}                         ".FormatTo(FirstConditionList.JoinWith(" AND ")) & _
                         "                        )                                     " & _
                         "             {0}                                              ".FormatTo(SecondConditionList.JoinWith(" "))


                If TableHasIdentity(FkItem.DatabaseName, FkItem.TableName) Then

                    StrSql = "SET IDENTITY_INSERT {0}..{1} ON ".FormatTo(FkItem.DatabaseName.Squared, StrNewTable) & vbNewLine & _
                             StrSql & vbNewLine & _
                             "SET IDENTITY_INSERT {0}..{1} OFF ".FormatTo(FkItem.DatabaseName.Squared, StrNewTable)

                End If


                StrSqlList.Add(StrSql)


                'Clipboard.Clear()
                'Clipboard.SetText(StrSql)


                'Gerando comando que eliminará os registros órfãos da tabela principal:
                SelectFieldList = GetPrimaryKeyFields(FkItem.DatabaseName, FkItem.TableName)    '--> obtendo campos chave da tabela


                StrSql = " DELETE FROM {0}..{1}                                 ".FormatTo(FkItem.DatabaseName.Squared, FkItem.TableName.Squared) & _
                         "      FROM   {0}..{1} AS [DependentTbl]               ".FormatTo(FkItem.DatabaseName.Squared, FkItem.TableName.Squared) & _
                         "      WHERE  NOT EXISTS (                             " & _
                         "                         SELECT   1                           " & _
                         "                         FROM     {0}..{1} AS [RelatedTbl]    ".FormatTo(FkItem.DatabaseName.Squared, FkItem.RelatedTableName.Squared) & _
                         "                         WHERE    {0}                         ".FormatTo(FirstConditionList.JoinWith(" AND ")) & _
                         "                        )                                     " & _
                         "            {0}   ".FormatTo(SecondConditionList.JoinWith(""))
                StrSqlList.Add(StrSql)


                'StrSql = " DELETE FROM {0}..{1}" & _
                '         "      FROM {0}..{1} AS [TBL]" & _
                '         "      WHERE EXISTS (" & _
                '         "          SELECT {0}                                      ".FormatTo(SelectFieldList.JoinWith(", ")) & _
                '         "          FROM   {0}..{1} AS [DependentTbl]               ".FormatTo(FkItem.DatabaseName.Squared, FkItem.TableName.Squared) & _
                '         "          WHERE  NOT EXISTS (                             " & _
                '         "                              SELECT  1                   " & _
                '         "                               FROM   {0}..{1} AS [RelatedTbl]    ".FormatTo(FkItem.DatabaseName.Squared, FkItem.RelatedTableName.Squared) & _
                '         "                              WHERE   {0}                         ".FormatTo(FirstConditionList.JoinWith(" AND ")) & _
                '         "                            )                                     " & _
                '         "                {0}   " & _
                '         "          )"
                'StrSqlList.Add(StrSql)


                'Executando os comandos na sequência:
                Me.Connection.BeginTransaction()

                For Each xSql In StrSqlList

                    'Clipboard.Clear()
                    'Clipboard.SetText(xSql)
                    Me.Connection.ExecuteNonQuery(xSql)
                Next

                Me.Connection.CommitTransaction()

                OrfanTableList.Add(StrNewTable)


                'Else
                'RaiseEvent LogEvent(Me, New LogEventArgs("OrfanRecordFix", "No orfan records were found"))
            End If


        Next


        If Not HasOrfanRecords Then
            RaiseEvent LogEvent(Me, New LogEventArgs("OrfanRecordFix", "No orfan records were found"))
        End If



        Return OrfanTableList
    End Function


    Private Shadows Sub VerificarSeUsuarioCancelou()
        If Not Me.BackgroundWorker.IsNull AndAlso Me.BackgroundWorker.CancellationPending Then
            Throw New Exception("Procedimento cancelado pelo usuário")
        End If
    End Sub



End Class

