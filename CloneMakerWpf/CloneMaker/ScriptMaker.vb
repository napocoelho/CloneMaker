Imports System.Data



Public Class ScriptMaker


    Public Property BackgroundWorker As System.ComponentModel.BackgroundWorker

    Public Event LogEvent As LogEventHandler


    Public Class DbFileInfo
        Public Property Name As String
        Public Property PhysicalName As String
        Public Property MaxSize As Long
        Public Property Growth As Long
        Public Property Size As Long
        Public Property SizeKB As Long

        Public Sub New()
            Name = String.Empty
            PhysicalName = String.Empty
            MaxSize = 0
            Growth = 0
            Size = 0
            SizeKB = 0
        End Sub
    End Class

    Public Class FieldInfo
        Public Property Name As String
        Public Property Type As String
        Public Property Precision As Integer
        Public Property Scale As Integer
        Public Property IsNullable As Boolean
        Public Property IsComputed As Boolean
        Public Property Collation As String
        Public Property IsIdentity As Boolean
        Public Property SeedValue As Long
        Public Property IncrementValue As Long
        Public Property IsPersisted As Boolean
        Public Property ComputedDefinition As String

        Public Sub New()
            Name = String.Empty
            Type = String.Empty
            Precision = 0
            Scale = 0
            IsNullable = False
            IsComputed = False
            Collation = String.Empty
            IsIdentity = False
            SeedValue = 0
            IncrementValue = 0
            IsPersisted = False
            ComputedDefinition = String.Empty
        End Sub
    End Class

    Public Enum CommandType
        Database
        Table
        DefaultConstraint
        CheckConstraint
        UniqueConstraint
        PrimaryKey
        ForeignKey
        StoredProcedure
        Trigger
        View
    End Enum
    



    'Private Function FormatField(ByVal FieldName As String, ByVal TypeName As String, _
    '                              ByVal lPrecision As Long, ByVal lScale As Long, _
    '                              ByVal IsNullable As Boolean, _
    '                              ByVal IsComputed As Boolean, _
    '                              Optional ByVal CollationName As String = "", _
    '                              Optional ByVal IsIdentity As Boolean = False, _
    '                              Optional ByVal SeedValue As Long = 0, _
    '                              Optional ByVal IncrementValue As Long = 0, _
    '                              Optional ByVal IsPersisted As Boolean = False, _
    '                              Optional ByVal StrComputedDefinition As String = "") As String



    Private ConnHlp As ConnectionManager

    Private Enum FileType
        MDF = 0
        LDF = 1
    End Enum

    Public Sub New(ByRef ConnArg As ConnectionManager)
        Me.BackgroundWorker = Nothing
        ConnHlp = ConnArg
    End Sub

    Public ReadOnly Property Connection As ConnectionManager
        Get
            Return Me.ConnHlp
        End Get
    End Property

    Private Function GetDbFileInfo(ByVal FileTypeArg As FileType) As DbFileInfo

        Dim StrSql As String, RsScripts As New DataTable
        Dim dbInfo As New DbFileInfo

        'GetDbFileInfo = New List(Of DbFileInfo)

        StrSql = " SELECT *, (size*8) AS size_KB FROM sys.database_files WHERE type = " & FileTypeArg

        RsScripts = ConnectionManager.GetInstance.ExecuteDataTable(StrSql)
        'RsScripts.Open(StrSql, Con, adOpenStatic, adLockReadOnly)

        'Terá apenas 1 linha:
        For Each xRow As DataRow In RsScripts.Rows
            dbInfo.Name = xRow("Name")
            dbInfo.PhysicalName = xRow("Physical_Name")
            dbInfo.MaxSize = xRow("Max_Size")
            dbInfo.Growth = xRow("Growth")
            dbInfo.Size = xRow("Size")
            dbInfo.SizeKB = xRow("Size_KB")
        Next

        GetDbFileInfo = dbInfo

    End Function



    Private Function FormatField(ByRef Field As FieldInfo) As String

        'Private Function FormatField(ByVal FieldName As String, ByVal TypeName As String, _
        '                              ByVal lPrecision As Long, ByVal lScale As Long, _
        '                              ByVal IsNullable As Boolean, _
        '                              ByVal IsComputed As Boolean, _
        '                              Optional ByVal CollationName As String = "", _
        '                              Optional ByVal IsIdentity As Boolean = False, _
        '                              Optional ByVal SeedValue As Long = 0, _
        '                              Optional ByVal IncrementValue As Long = 0, _
        '                              Optional ByVal IsPersisted As Boolean = False, _
        '                              Optional ByVal StrComputedDefinition As String = "") As String

        Dim DeclarationList As New List(Of String)


        Field.Name = Field.Name.Trim
        Field.Type = Field.Type.ToLower



        If Field.IsComputed Then

            'Declaração de Nome:
            DeclarationList.Add(Field.Name.Squared)
            DeclarationList.Add("AS")
            DeclarationList.Add(Field.ComputedDefinition)

            If Field.IsPersisted Then

                DeclarationList.Add("PERSISTED")

                If Not Field.IsNullable Then _
                   DeclarationList.Add("NOT NULL")

            End If

        Else

            'Declaração de Nome e Tipo:
            DeclarationList.Add(Field.Name.Squared)
            DeclarationList.Add(Field.Type.Squared)



            'Não possui [Precision] nem [Scale]
            If Field.Type.ExistsIn("int", "bit", "money", "text", "ntext", "tinyint", "smallint", "bigint", "smallmoney", "real", "datetime", "datetime2", "smalldatetime") Then

                '---> Não possui [Precision] nem [Scale] <---'

                'Possui [Precision], mas não [Scale]
            ElseIf Field.Type.ExistsIn("varchar", "nvarchar", "char", "nchar", "binary", "float", "time", "datetimeoffset") Then

                DeclarationList.Add(Field.Precision.ToString.Bracket)

            ElseIf Field.Type.ExistsIn("varbinary") Then

                DeclarationList.Add(IIf(Field.Precision > 0, Field.Precision, "max").ToString.Bracket)

                'Possui [Precision] e [Scale]
            ElseIf Field.Type.ExistsIn("decimal", "numeric") Then

                DeclarationList.Add((Field.Precision & ", " & Field.Scale).Bracket)

                'Versionados:
            ElseIf Field.Type.ExistsIn("timestamp", "rowversion", "uniqueidentifier") Then

                '---> Não possui [Precision] nem [Scale] <---'

            End If


            'Se possuir COLLATION:
            '*** OBS.: Deve ficar logo após a declaração do tipo ***
            If Not Field.Collation.IsEmpty Then

                DeclarationList.Add("COLLATE " & Field.Collation)
            End If

            'Campo anulável (OU NÃO):
            DeclarationList.Add(IIf(Field.IsNullable, "NULL", "NOT NULL"))


            If Field.IsIdentity Then

                DeclarationList.Add("IDENTITY (" & Field.SeedValue & ", " & Field.IncrementValue & ")")
            End If

        End If


        'Unindo todas as configurações:
        FormatField = vbTab & DeclarationList.JoinWith(Space(1))

        DeclarationList = Nothing
    End Function

    Private Function GetOnOff(ByVal YesOrNo As Boolean) As String
        GetOnOff = IIf(YesOrNo, "ON", "OFF")
    End Function

    Public Function MakeDatabase(ByVal DbName As String, Optional ByVal NewDbName As String = "") As List(Of String)

        Dim DeclarationList As New List(Of String)
        Dim StrSql As String
        Dim RsDb As DataTable

        If NewDbName.IsEmpty Then
            NewDbName = DbName
        End If

        RaiseEvent LogEvent(Me, New LogEventArgs("MakeDatabase", "Creating script from [{0}] to [{0}]".FormatTo(DbName, NewDbName)))

        StrSql = " SELECT * FROM " & GetViewSchemaDatabases & " WHERE name = " & Quote(DbName) & " ORDER BY Name"

        RsDb = Connection.ExecuteDataTable(StrSql)


        For Each xRow As DataRow In RsDb.Rows

            Dim StrDeclaration As String, DbInfo As New DbFileInfo
            Dim StrMdfDeclaration As String, StrLdfDeclaration As String


            'Criando declaração formatada:
            '--------------------------------------------------------------------------------------------------------------------------------------------
            DbInfo = GetDbFileInfo(FileType.MDF)
            StrMdfDeclaration = vbTab & "( NAME = N{0}, FILENAME = N{1}, SIZE = {2}KB, MAXSIZE = {3}, FILEGROWTH = {4}% )" & vbNewLine
            StrMdfDeclaration = StrMdfDeclaration.FormatTo(DbInfo.Name.Quote, _
                                                           DbInfo.PhysicalName.Quote, _
                                                           5120, _
                                                           IIf(DbInfo.MaxSize = -1, "UNLIMITED", DbInfo.MaxSize & "KB"), _
                                                           DbInfo.Growth)

            DbInfo = GetDbFileInfo(FileType.LDF)
            StrLdfDeclaration = vbTab & "( NAME = N{0}, FILENAME = N{1}, SIZE = {2}KB, MAXSIZE = {3}, FILEGROWTH = {4}% )"
            StrLdfDeclaration = StrLdfDeclaration.FormatTo(DbInfo.Name.Quote, _
                                                           DbInfo.PhysicalName.Quote, _
                                                           512, _
                                                           IIf(DbInfo.MaxSize = -1, "UNLIMITED", DbInfo.MaxSize & "KB"), _
                                                           DbInfo.Growth)


            'Aqui, usa-se o [DbName] mesmo (ao invés do [NewDbName], pois será renomeado logo abaixo:
            StrDeclaration = "CREATE DATABASE " & DbName.Squared & " ON PRIMARY" & vbNewLine & _
                              StrMdfDeclaration & _
                              vbTab & " LOG ON" & vbNewLine & _
                              StrLdfDeclaration

            StrDeclaration = Microsoft.VisualBasic.Strings.Replace(StrDeclaration, DbName, NewDbName, , , CompareMethod.Text) '---> importante (comparação case-insensitive)

            '--------------------------------------------------------------------------------------------------------------------------------------------


            DeclarationList.Add(StrDeclaration)


            'Scripts de configuração da nova base:

            DeclarationList.Add("ALTER DATABASE " & NewDbName.Squared & " SET COMPATIBILITY_LEVEL = " & xRow("Compatibility_Level"))

            DeclarationList.Add("IF (1 = FULLTEXTSERVICEPROPERTY('IsFullTextInstalled'))" & vbNewLine & _
                                 "begin" & vbNewLine & _
                                 "EXEC [" & NewDbName & "].[dbo].[sp_fulltext_database] @action = 'disable'" & vbNewLine & _
                                 "end")

            DeclarationList.Add("ALTER DATABASE " & NewDbName.Squared & " COLLATE " & xRow("collation_name"))

            DeclarationList.Add("ALTER DATABASE " & NewDbName.Squared & " SET ANSI_NULL_DEFAULT " & GetOnOff(xRow("Is_Ansi_Null_Default_On")))

            DeclarationList.Add("ALTER DATABASE " & NewDbName.Squared & " SET ANSI_NULLS " & GetOnOff(xRow("Is_Ansi_Nulls_On")))

            DeclarationList.Add("ALTER DATABASE " & NewDbName.Squared & " SET ANSI_WARNINGS " & GetOnOff(xRow("Is_Ansi_Warnings_On")))

            DeclarationList.Add("ALTER DATABASE " & NewDbName.Squared & " SET ARITHABORT " & GetOnOff(xRow("Is_ArithAbort_On")))

            DeclarationList.Add("ALTER DATABASE " & NewDbName.Squared & " SET AUTO_CLOSE " & GetOnOff(xRow("Is_Auto_Close_On")))

            DeclarationList.Add("ALTER DATABASE " & NewDbName.Squared & " SET AUTO_CREATE_STATISTICS " & GetOnOff(xRow("Is_Auto_Create_Stats_On")))

            DeclarationList.Add("ALTER DATABASE " & NewDbName.Squared & " SET AUTO_SHRINK " & GetOnOff(xRow("Is_Auto_Shrink_On")))

            DeclarationList.Add("ALTER DATABASE " & NewDbName.Squared & " SET AUTO_UPDATE_STATISTICS " & GetOnOff(xRow("Is_Auto_Update_Stats_On")))

            DeclarationList.Add("ALTER DATABASE " & NewDbName.Squared & " SET CURSOR_CLOSE_ON_COMMIT " & GetOnOff(xRow("Is_Cursor_Close_On_Commit_On")))

            DeclarationList.Add("ALTER DATABASE " & NewDbName.Squared & " SET CURSOR_DEFAULT " & IIf(xRow("Is_Local_Cursor_Default") = 0, "GLOBAL", "LOCAL"))

            DeclarationList.Add("ALTER DATABASE " & NewDbName.Squared & " SET CONCAT_NULL_YIELDS_NULL " & GetOnOff(xRow("Is_Concat_Null_Yields_Null_On")))

            DeclarationList.Add("ALTER DATABASE " & NewDbName.Squared & " SET NUMERIC_ROUNDABORT " & GetOnOff(xRow("Is_Numeric_RoundAbort_On")))

            DeclarationList.Add("ALTER DATABASE " & NewDbName.Squared & " SET QUOTED_IDENTIFIER " & GetOnOff(xRow("Is_Quoted_Identifier_On")))

            DeclarationList.Add("ALTER DATABASE " & NewDbName.Squared & " SET RECURSIVE_TRIGGERS " & GetOnOff(xRow("Is_Recursive_Triggers_On")))

            DeclarationList.Add("ALTER DATABASE " & NewDbName.Squared & " SET " & IIf(xRow("Is_Broker_Enabled"), "ENABLE_BROKER", "DISABLE_BROKER"))

            DeclarationList.Add("ALTER DATABASE " & NewDbName.Squared & " SET AUTO_UPDATE_STATISTICS_ASYNC " & GetOnOff(xRow("Is_Auto_Update_Stats_Async_On")))

            DeclarationList.Add("ALTER DATABASE " & NewDbName.Squared & " SET DATE_CORRELATION_OPTIMIZATION " & GetOnOff(xRow("Is_Date_Correlation_On")))

            DeclarationList.Add("ALTER DATABASE " & NewDbName.Squared & " SET TRUSTWORTHY " & GetOnOff(xRow("Is_TrustWorthy_On")))

            DeclarationList.Add("ALTER DATABASE " & NewDbName.Squared & " SET ALLOW_SNAPSHOT_ISOLATION " & GetOnOff(xRow("SnapShot_Isolation_State")))

            DeclarationList.Add("ALTER DATABASE " & NewDbName.Squared & " SET PARAMETERIZATION " & IIf(xRow("Is_Parameterization_Forced"), "FORCED", "SIMPLE"))

            DeclarationList.Add("ALTER DATABASE " & NewDbName.Squared & " SET READ_COMMITTED_SNAPSHOT " & GetOnOff(xRow("Is_Read_Committed_Snapshot_On")))

            DeclarationList.Add("ALTER DATABASE " & NewDbName.Squared & " SET HONOR_BROKER_PRIORITY " & GetOnOff(xRow("Is_Honor_Broker_Priority_On")))

            DeclarationList.Add("ALTER DATABASE " & NewDbName.Squared & " SET RECOVERY " & xRow("Recovery_Model_Desc"))

            DeclarationList.Add("ALTER DATABASE " & NewDbName.Squared & " SET " & xRow("User_Access_Desc"))

            DeclarationList.Add("ALTER DATABASE " & NewDbName.Squared & " SET PAGE_VERIFY " & xRow("Page_Verify_Option_Desc"))

            DeclarationList.Add("ALTER DATABASE " & NewDbName.Squared & " SET DB_CHAINING " & GetOnOff(xRow("Is_Db_Chaining_On")))

            DeclarationList.Add("ALTER DATABASE " & NewDbName.Squared & " SET " & IIf(xRow("Is_Read_Only"), "READ_ONLY", "READ_WRITE"))


            'Clipboard.Clear
            'Clipboard.SetText (StrDeclaration)

        Next



        'MakeDatabase = DeclarationList

        'MakeDatabase = "/**** CREATING DATABASE STRUCTURE ****/" & vbNewLine & _
        '          DeclarationList.JoinWith(vbNewLine & "GO" & vbNewLine & vbNewLine) & _
        '          vbNewLine & "GO" & vbNewLine & vbNewLine

        Return DeclarationList
    End Function

    Public Function MakeTables(Optional ByVal TableName As String = "") As List(Of String)

        Dim StrSql As String, StrWhere As String
        Dim RsTables As DataTable
        Dim RsFields As DataTable

        Dim ListFieldsDeclaration As List(Of String)
        Dim ListTablesDeclaration As List(Of String)





        StrWhere = IIf(Not TableName.IsEmpty, " WHERE TABLE_NAME = " & TableName.Quote, "")

        'If Not TableName.IsEmpty Then

        '    StrWhere = " WHERE TABLE_NAME = " & TableName.Quote
        'End If

        StrSql = " SELECT DISTINCT TABLE_NAME FROM " & GetViewSchemaAllColumns & Space(1) & StrWhere

        RsTables = Me.Connection.ExecuteDataTable(StrSql)

        ListTablesDeclaration = New List(Of String)


        For Each xRow As DataRow In RsTables.Rows

            Dim StrFieldsDeclaration As String


            VerificarSeUsuarioCancelou()
            RaiseEvent LogEvent(Me, New LogEventArgs("MakeTables", "Scripting table [{0}]".FormatTo(xRow("TABLE_NAME").ToString)))



            StrSql = " SELECT VIEW_SCHEMA_ALL_COLUMNS.Table_Name, VIEW_SCHEMA_ALL_COLUMNS.Column_Name, Type_Name, Precision, Scale, IsNullable, IsComputed, IsPersisted, ComputedDefinition, Collation, " & _
                     "        IsIdentity, Seed_Value, Increment_Value" & _
                     " FROM   " & GetViewSchemaAllColumns & _
                     "        LEFT JOIN " & GetViewSchemaConstraintId & _
                     "           ON VIEW_CONSTRAINTS_ID.Table_Name = VIEW_SCHEMA_ALL_COLUMNS.Table_Name " & _
                     "           AND VIEW_CONSTRAINTS_ID.Column_Name = VIEW_SCHEMA_ALL_COLUMNS.Column_Name " & _
                     " WHERE  VIEW_SCHEMA_ALL_COLUMNS.TABLE_NAME = " & Quote(xRow("TABLE_NAME")) & _
                     " ORDER  BY VIEW_SCHEMA_ALL_COLUMNS.TABLE_NAME, COLUMN_ORDER"
            RsFields = Connection.ExecuteDataTable(StrSql)

            ListFieldsDeclaration = New List(Of String)


            For Each xRow2 As DataRow In RsFields.Rows

                Dim StrDeclaration As String
                Dim NewField As FieldInfo

                NewField = New FieldInfo
                NewField.Name = xRow2("COLUMN_NAME")
                NewField.Type = xRow2("TYPE_NAME")
                NewField.Precision = xRow2("PRECISION").ToString.IfIsEmpty(0)
                NewField.Scale = xRow2("SCALE").ToString.IfIsEmpty(0)
                NewField.IsNullable = xRow2("ISNULLABLE")
                NewField.IsComputed = xRow2("ISCOMPUTED")
                NewField.Collation = xRow2("COLLATION").ToString
                NewField.IsIdentity = xRow2("IsIdentity")
                NewField.SeedValue = xRow2("Seed_Value").ToString.IfIsEmpty(0)
                NewField.IncrementValue = xRow2("Increment_Value").ToString.IfIsEmpty(0)
                NewField.IsPersisted = xRow2("IsPersisted")
                NewField.ComputedDefinition = xRow2("ComputedDefinition").ToString

                StrDeclaration = FormatField(NewField)

                ListFieldsDeclaration.Add(StrDeclaration)
            Next




            StrFieldsDeclaration = ListFieldsDeclaration.JoinWith(", " & vbNewLine)


            StrFieldsDeclaration = "CREATE TABLE " & xRow("TABLE_NAME").ToString.Squared & vbNewLine & _
                                   "(" & vbNewLine & _
                                   StrFieldsDeclaration & vbNewLine & _
                                   ")"

            'Clipboard.Clear
            'Clipboard.SetText (StrFieldsDeclaration)


            ListTablesDeclaration.Add(StrFieldsDeclaration)

        Next





        'MakeTables = "/**** CREATING TABLES AND FIELDS STRUCTURE ****/" & vbNewLine & _
        '        ListTablesDeclaration.JoinItems(vbNewLine & "GO" & vbNewLine & vbNewLine) & _
        '          vbNewLine & "GO" & vbNewLine & vbNewLine


        'Clipboard.Clear
        'Clipboard.SetText (MakeTables)


        Return ListTablesDeclaration
    End Function



    Public Function MakeDefaultConstraints() As List(Of String)

        Dim ListDfDeclaration As New List(Of String)
        Dim StrSql As String
        Dim RsDF As DataTable

        StrSql = " SELECT Table_Name, Constraint_Name, Column_Name, Definition FROM " & GetViewSchemaConstraintDf & " ORDER BY Table_Name, Column_Name"

        RsDF = Connection.ExecuteDataTable(StrSql)


        For Each xRow As DataRow In RsDF.Rows
            Dim StrDeclaration As String, StrDefinition As String

            VerificarSeUsuarioCancelou()
            RaiseEvent LogEvent(Me, New LogEventArgs("MakeDefaultConstraints", "Scripting [{0}] : [{1}].[{2}]".FormatTo(xRow("Constraint_Name").ToString, xRow("Table_Name").ToString, xRow("Column_Name").ToString)))


            StrDefinition = xRow("Definition")

            'Eliminando lixo:
            If Left(StrDefinition, 1) = "(" And Right(StrDefinition, 1) = ")" Then
                StrDefinition = Left(StrDefinition, Len(StrDefinition) - 1)
                StrDefinition = Right(StrDefinition, Len(StrDefinition) - 1)
            End If

            'Criando declaração formatada:
            StrDeclaration = "ALTER TABLE {0}" & vbNewLine & _
                              vbTab & "ADD CONSTRAINT {1}" & vbNewLine & _
                              vbTab & vbTab & "DEFAULT {2} FOR {3}"

            StrDeclaration = StrDeclaration.FormatTo( _
                                             xRow("Table_Name").ToString.Squared, _
                                             xRow("Constraint_Name").ToString.Squared, _
                                             StrDefinition, _
                                             xRow("Column_Name").ToString.Squared)

            ListDfDeclaration.Add(StrDeclaration)

            'Clipboard.Clear
            'Clipboard.SetText (StrDeclaration)
        Next



        'MakeDefaultConstraints = "/**** CREATING DEFAULT CONSTRAINTS STRUCTURE ****/" & vbNewLine & _
        '                   ListDfDeclaration.JoinItems(vbNewLine & "GO" & vbNewLine & vbNewLine) & _
        '                     vbNewLine & "GO" & vbNewLine & vbNewLine

        'Clipboard.Clear
        'Clipboard.SetText (MakeDefaultConstraints)


        Return ListDfDeclaration
    End Function

    Public Function MakeCheckConstraints() As List(Of String)

        Dim ListDfDeclaration As New List(Of String)
        Dim StrSql As String
        Dim RsDF As DataTable

        StrSql = " SELECT Table_Name, Constraint_Name, Column_Name, Definition FROM " & GetViewSchemaConstraintCk & " ORDER BY Table_Name, Column_Name"

        RsDF = Connection.ExecuteDataTable(StrSql)

        For Each xRow As DataRow In RsDF.Rows

            Dim StrDeclaration As String, StrDefinition As String


            VerificarSeUsuarioCancelou()
            RaiseEvent LogEvent(Me, New LogEventArgs("MakeCheckConstraints", "Scripting [{0}] : [{1}].[{2}]".FormatTo(xRow("Constraint_Name").ToString, xRow("Table_Name").ToString, xRow("Column_Name").ToString)))

            StrDefinition = xRow("Definition")

            'Eliminando lixo:
            If Left(StrDefinition, 1) = "(" And Right(StrDefinition, 1) = ")" Then
                StrDefinition = Left(StrDefinition, Len(StrDefinition) - 1)
                StrDefinition = Right(StrDefinition, Len(StrDefinition) - 1)
            End If

            'Criando declaração formatada:
            StrDeclaration = "ALTER TABLE {0}" & vbNewLine & _
                              vbTab & "ADD CONSTRAINT {1}" & vbNewLine & _
                              vbTab & vbTab & "CHECK ( {2} )"

            StrDeclaration = StrDeclaration.FormatTo( _
                                             xRow("Table_Name").ToString.Squared, _
                                             xRow("Constraint_Name").ToString.Squared, _
                                             StrDefinition)

            ListDfDeclaration.Add(StrDeclaration)

            'Clipboard.Clear
            'Clipboard.SetText (StrDeclaration)

        Next





        'MakeCheckConstraints = "/**** CREATING DEFAULT CONSTRAINTS STRUCTURE ****/" & vbNewLine & _
        '                    ListDfDeclaration.JoinItems(vbNewLine & "GO" & vbNewLine & vbNewLine) & _
        '                   vbNewLine & "GO" & vbNewLine & vbNewLine

        'Clipboard.Clear
        'Clipboard.SetText (MakeCheckConstraints)


        Return ListDfDeclaration
    End Function

    Public Function MakeUniqueConstraints() As List(Of String)

        Dim StrSql As String
        Dim StrDeclaration As String
        Dim RsConstraints As DataTable
        Dim RsFields As DataTable

        Dim ListFields As List(Of String)
        Dim ListConstraintDeclaration As List(Of String)

        StrSql = " SELECT DISTINCT Table_Name, Pk_Name FROM " & GetViewSchemaConstraintUq & " ORDER BY Table_Name, Pk_Name"

        RsConstraints = Connection.ExecuteDataTable(StrSql)

        ListConstraintDeclaration = New List(Of String)

        For Each xRow As DataRow In RsConstraints.Rows

            Dim StrFields As String

            VerificarSeUsuarioCancelou()
            RaiseEvent LogEvent(Me, New LogEventArgs("MakeUniqueConstraints", "Scripting [{0}] : [{1}]".FormatTo(xRow("Pk_Name").ToString, xRow("Table_Name").ToString)))


            StrSql = " SELECT DISTINCT Table_Name, Pk_Name, Column_Name " & _
                     " FROM   " & GetViewSchemaConstraintUq & _
                     " WHERE  Table_Name = " & Quote(xRow("Table_Name")) & _
                     "        AND Pk_Name = " & Quote(xRow("Pk_Name")) & _
                     " ORDER  BY Table_Name, Pk_Name, Column_Name"
            RsFields = Me.Connection.ExecuteDataTable(StrSql)

            ListFields = New List(Of String)


            For Each xRow2 As DataRow In RsFields.Rows
                ListFields.Add(xRow2("Column_Name").ToString.Squared)
            Next





            StrFields = ListFields.JoinWith(", ")


            StrDeclaration = "ALTER TABLE {0}" & vbNewLine & _
                                   vbTab & "ADD CONSTRAINT {1}" & vbNewLine & _
                                   vbTab & vbTab & "UNIQUE ( {2} )"

            StrDeclaration = StrDeclaration.FormatTo( _
                                          xRow("Table_Name").ToString.Squared, _
                                          xRow("Pk_Name").ToString.Squared, _
                                          StrFields)

            'Clipboard.Clear
            'Clipboard.SetText (StrFieldsDeclaration)

            ListConstraintDeclaration.Add(StrDeclaration)

        Next








        'MakeUniqueConstraints = "/**** CREATING UNIQUE CONSTRAINTS STRUCTURE ****/" & vbNewLine & _
        '                   ListConstraintDeclaration.JoinItems(vbNewLine & "GO" & vbNewLine & vbNewLine) & _
        '             vbNewLine & "GO" & vbNewLine & vbNewLine


        'Clipboard.Clear
        'Clipboard.SetText (MakeUniqueConstraints)


        Return ListConstraintDeclaration
    End Function


    Public Function MakePrimaryKeys() As List(Of String)

        Dim StrSql As String
        Dim StrDeclaration As String
        Dim RsConstraints As DataTable
        Dim RsFields As DataTable

        Dim ListFields As List(Of String)
        Dim ListConstraintDeclaration As List(Of String)

        StrSql = " SELECT DISTINCT TableName, PkName FROM " & GetViewSchemaConstraintPk & " ORDER BY TableName, PkName"

        RsConstraints = Connection.ExecuteDataTable(StrSql)

        ListConstraintDeclaration = New List(Of String)


        For Each xRow In RsConstraints.Rows

            Dim StrFields As String


            VerificarSeUsuarioCancelou()
            RaiseEvent LogEvent(Me, New LogEventArgs("MakePrimaryKeys", "Scripting [{0}] : [{1}]".FormatTo(xRow("PkName").ToString, xRow("TableName").ToString)))


            '      If xRow("TableName") = "TBL_BANCOS" Then
            '         Stop
            '      End If

            StrSql = " SELECT DISTINCT TableName, PkName, ColumnName, ColumnOrder " & _
                     " FROM   " & GetViewSchemaConstraintPk & _
                     " WHERE  TableName = " & xRow("TableName").ToString.Quote & _
                     "        AND PkName = " & xRow("PkName").ToString.Quote & _
                     " ORDER  BY TableName, PkName, ColumnOrder"

            RsFields = Connection.ExecuteDataTable(StrSql)

            ListFields = New List(Of String)




            For Each xRow2 In RsFields.Rows
                ListFields.Add(xRow2("ColumnName").ToString.Squared)
            Next



            StrFields = ListFields.JoinWith(", ")


            StrDeclaration = "ALTER TABLE {0}" & vbNewLine & _
                                   vbTab & "ADD CONSTRAINT {1}" & vbNewLine & _
                                   vbTab & vbTab & "PRIMARY KEY ( {2} )"

            StrDeclaration = StrDeclaration.FormatTo( _
                                          xRow("TableName").ToString.Squared, _
                                          xRow("PkName").ToString.Squared, _
                                          StrFields)

            'Clipboard.Clear
            'Clipboard.SetText (StrFieldsDeclaration)

            ListConstraintDeclaration.Add(StrDeclaration)

        Next







        'MakePrimaryKeys = "/**** CREATING PRIMARY KEY CONSTRAINTS STRUCTURE ****/" & vbNewLine & _
        '             ListConstraintDeclaration.JoinItems(vbNewLine & "GO" & vbNewLine & vbNewLine) & _
        '             vbNewLine & "GO" & vbNewLine & vbNewLine


        'Clipboard.Clear
        'Clipboard.SetText (MakePrimaryKeys)

        Return ListConstraintDeclaration
    End Function

    Public Function MakeForeignKeys() As List(Of String)

        Dim StrSql As String
        Dim StrDeclaration As String
        Dim RsConstraints As DataTable
        Dim RsFields As DataTable

        Dim ListSourceFields As List(Of String), ListDependentFields As List(Of String)
        Dim ListConstraintDeclaration As List(Of String)

        StrSql = " SELECT DISTINCT Fk_Name, SourceTable, DependentTable FROM " & GetViewSchemaConstraintFk & " ORDER BY Fk_Name "

        RsConstraints = Connection.ExecuteDataTable(StrSql)

        ListConstraintDeclaration = New List(Of String)


        'Obtaining correlated fields:
        For Each xRow In RsConstraints.Rows

            Dim StrSourceFields As String, StrDependentFields As String, StrOnDelete As String, StrOnUpdate As String
            Dim BooOnFirstTime As Boolean = False

            StrOnDelete = String.Empty
            StrOnUpdate = String.Empty

            VerificarSeUsuarioCancelou()
            RaiseEvent LogEvent(Me, New LogEventArgs("MakeForeignKeys", "Scripting [{0}] : [{1}] belongs to [{2}]".FormatTo(xRow("Fk_Name").ToString, xRow("DependentTable").ToString, xRow("SourceTable").ToString)))

            StrSql = " SELECT Fk_Name, SourceTable, DependentTable, SourceOrder, SourceCol, DependentOrder, DependentCol, Update_Rule, Delete_Rule, Match_Option " & _
                     " FROM   " & GetViewSchemaConstraintFk & _
                     " WHERE  Fk_Name = " & Quote(xRow("Fk_Name")) & _
                     " ORDER  BY Fk_Name, SourceTable, DependentTable, SourceOrder, SourceCol, DependentOrder, DependentCol"
            RsFields = Connection.ExecuteDataTable(StrSql)

            ListSourceFields = New List(Of String)
            ListDependentFields = New List(Of String)

            For Each xRow2 In RsFields.Rows
                ListSourceFields.Add(xRow2("SourceCol").ToString.Squared)
                ListDependentFields.Add(xRow2("DependentCol").ToString.Squared)

                If Not BooOnFirstTime Then
                    BooOnFirstTime = True

                    StrOnDelete = xRow2("Delete_Rule").ToString.ToUpper
                    StrOnDelete = IIf(StrOnDelete = "NO ACTION", String.Empty, "ON DELETE " & StrOnDelete)

                    StrOnUpdate = xRow2("Update_Rule").ToString.ToUpper
                    StrOnUpdate = IIf(StrOnUpdate = "NO ACTION", String.Empty, "ON UPDATE " & StrOnUpdate)
                End If
            Next


            StrSourceFields = ListSourceFields.JoinWith(", ")
            StrDependentFields = ListDependentFields.JoinWith(", ")

            StrDeclaration = "ALTER TABLE {0}" & vbNewLine & _
                                   vbTab & "ADD CONSTRAINT {1}" & vbNewLine & _
                                   vbTab & vbTab & "FOREIGN KEY ( {2} )" & vbNewLine & _
                                   vbTab & vbTab & vbTab & "REFERENCES {3} ( {4} )" & vbNewLine & _
                                   vbTab & vbTab & vbTab & vbTab & " {5} {6}"

            StrDeclaration = StrDeclaration.FormatTo( _
                                          xRow("DependentTable").ToString.Squared, _
                                          xRow("Fk_Name").ToString.Squared, _
                                          StrDependentFields, _
                                          xRow("SourceTable").ToString.Squared, _
                                          StrSourceFields,
                                          StrOnDelete,
                                          StrOnUpdate)

            'Clipboard.Clear
            'Clipboard.SetText (StrFieldsDeclaration)

            ListConstraintDeclaration.Add(StrDeclaration)

        Next







        'MakeForeignKeys = "/**** CREATING FOREIGN KEY CONSTRAINTS STRUCTURE ****/" & vbNewLine & _
        '             ListConstraintDeclaration.JoinItems(vbNewLine & "GO" & vbNewLine & vbNewLine) & _
        '             vbNewLine & "GO" & vbNewLine & vbNewLine


        'Clipboard.Clear
        'Clipboard.SetText (MakeForeignKeys)

        Return ListConstraintDeclaration
    End Function

    Public Function MakeTriggers() As List(Of String)

        Dim StrSql As String
        Dim RsTriggers As DataTable
        Dim RsTriggersPart As DataTable

        Dim ListConstraintDeclaration As List(Of String)

        StrSql = " SELECT DISTINCT Fn_Name FROM " & GetViewSchemaTriggers & " ORDER BY Fn_Name "

        RsTriggers = Connection.ExecuteDataTable(StrSql)

        ListConstraintDeclaration = New List(Of String)


        'Obtaining correlated fields:
        For Each xRow As DataRow In RsTriggers.Rows
            Dim StrScript As String

            VerificarSeUsuarioCancelou()
            RaiseEvent LogEvent(Me, New LogEventArgs("MakeTriggers", "Scripting [{0}]".FormatTo(xRow("Fn_Name").ToString)))



            StrScript = String.Empty

            StrSql = " SELECT Fn_Name, ColId, Script " & _
                     " FROM   " & GetViewSchemaTriggers & _
                     " WHERE  Fn_Name = " & xRow("Fn_Name").ToString.Quote & _
                     " ORDER  BY Fn_Name, ColId "
            RsTriggersPart = Connection.ExecuteDataTable(StrSql)


            For Each xRow2 As DataRow In RsTriggersPart.Rows
                StrScript = StrScript & xRow2("Script")
            Next



            ListConstraintDeclaration.Add(StrScript.Trim)
        Next





        'MakeTriggers = "/**** CREATING TRIGGERS STRUCTURE ****/" & vbNewLine & _
        '             ListConstraintDeclaration.JoinItems(vbNewLine & "GO" & vbNewLine & vbNewLine) & _
        '             vbNewLine & "GO" & vbNewLine & vbNewLine


        'Clipboard.Clear
        'Clipboard.SetText (MakeForeignKeys)


        Return ListConstraintDeclaration
    End Function

    Public Function MakeFunctions() As List(Of String)

        Dim StrSql As String
        Dim RsFunctions As DataTable
        Dim RsFunctionsPart As DataTable
        Dim ListConstraintDeclaration As List(Of String)



        StrSql = " SELECT DISTINCT Fn_Name FROM " & GetViewSchemaFunctions & " ORDER BY Fn_Name "

        RsFunctions = Connection.ExecuteDataTable(StrSql)

        ListConstraintDeclaration = New List(Of String)


        'Obtaining correlated fields:
        For Each xRow As DataRow In RsFunctions.Rows


            VerificarSeUsuarioCancelou()

            Dim StrScript As String = String.Empty

            RaiseEvent LogEvent(Me, New LogEventArgs("MakeFunctions", "Scripting [{0}]".FormatTo(xRow("Fn_Name").ToString)))

            StrSql = " SELECT Fn_Name, ColId, Script " & _
                     " FROM   " & GetViewSchemaFunctions & _
                     " WHERE  Fn_Name = " & xRow("Fn_Name").ToString.Quote & _
                     " ORDER  BY Fn_Name, ColId "
            RsFunctionsPart = Connection.ExecuteDataTable(StrSql)

            For Each xRow2 As DataRow In RsFunctionsPart.Rows
                StrScript = StrScript & xRow2("Script").ToString
            Next

            ListConstraintDeclaration.Add(StrScript.Trim)

        Next







        'MakeFunctions = "/**** CREATING FUNCTIONS STRUCTURE ****/" & vbNewLine & _
        '             ListConstraintDeclaration.JoinItems(vbNewLine & "GO" & vbNewLine & vbNewLine) & _
        '             vbNewLine & "GO" & vbNewLine & vbNewLine


        'Clipboard.Clear
        'Clipboard.SetText (MakeForeignKeys)

        Return ListConstraintDeclaration
    End Function

    Public Function MakeViews() As List(Of String)

        Dim StrSql As String
        Dim RsViews As DataTable
        Dim RsViewsPart As DataTable

        Dim ListConstraintDeclaration As List(Of String)

        StrSql = " SELECT DISTINCT View_Name FROM " & GetViewSchemaViews & " ORDER BY View_Name "

        RsViews = Connection.ExecuteDataTable(StrSql)

        ListConstraintDeclaration = New List(Of String)

        'Obtaining correlated fields:
        For Each xRow As DataRow In RsViews.Rows

            VerificarSeUsuarioCancelou()

            Dim StrScript As String

            RaiseEvent LogEvent(Me, New LogEventArgs("MakeViews", "Scripting [{0}]".FormatTo(xRow("View_Name").ToString)))


            StrScript = String.Empty

            StrSql = " SELECT View_Name, ColId, Script " & _
                     " FROM   " & GetViewSchemaViews & _
                     " WHERE  View_Name = " & xRow("View_Name").ToString.Quote & _
                     " ORDER  BY View_Name, ColId "
            RsViewsPart = Connection.ExecuteDataTable(StrSql)

            For Each xRow2 As DataRow In RsViewsPart.Rows
                StrScript = StrScript & xRow2("Script").ToString
            Next

            ListConstraintDeclaration.Add(Trim(StrScript))

        Next







        'MakeViews = "/**** CREATING VIEWS STRUCTURE ****/" & vbNewLine & _
        '       ListConstraintDeclaration.JoinItems(vbNewLine & "GO" & vbNewLine & vbNewLine) & _
        '       vbNewLine & "GO" & vbNewLine & vbNewLine


        'Clipboard.Clear
        'Clipboard.SetText (MakeForeignKeys)

        Return ListConstraintDeclaration
    End Function

    Private ReadOnly Property GetViewSchemaAllColumns As String
        Get

            Dim StrSql As String

            StrSql = ""
            StrSql = StrSql & "SELECT SCHEMA_NAME(UID)             AS SCHEMA_NAME , "
            StrSql = StrSql & "       sysobjects.ID                AS TABLE_ID    , "
            StrSql = StrSql & "       sysobjects.name              AS TABLE_NAME  , "
            StrSql = StrSql & "       COLUMNS_1.name               AS COLUMN_NAME , "
            StrSql = StrSql & "       COLUMNS_1.COLID              AS COLUMN_ORDER, "
            StrSql = StrSql & "       TYPE_NAME(COLUMNS_1.xtype)   AS TYPE_NAME   , "
            StrSql = StrSql & "       Collation                                   , "
            StrSql = StrSql & "       CollationId                                 , "
            StrSql = StrSql & "       Prec AS PRECISION                           , "
            StrSql = StrSql & "       COLUMNS_1.Scale                             , "
            StrSql = StrSql & "       IsNullable                                  , "
            StrSql = StrSql & "       IsComputed                                  , "
            StrSql = StrSql & "       IsNull(c_is_persisted, 0) AS IsPersisted    , "
            StrSql = StrSql & "       c_definition AS ComputedDefinition          , "
            StrSql = StrSql & "       Is_RowGuidCol                  AS IsRowGuidCol               , "
            StrSql = StrSql & "       Is_Identity                    AS IsIdentity                 , "
            StrSql = StrSql & "       Is_FileStream                  AS IsFileStream               , "
            StrSql = StrSql & "       Is_Replicated                  AS IsReplicated               , "
            StrSql = StrSql & "       Is_Non_Sql_Subscribed          AS IsNonSqlSubscribed         , "
            StrSql = StrSql & "       Is_Ansi_Padded                 AS IsAnsiPadded               , "
            StrSql = StrSql & "       Is_Merge_Published             AS IsMergePublished           , "
            StrSql = StrSql & "       Default_Object_Id              AS DF_CONSTRAINT_ID           , "
            StrSql = StrSql & "       OBJECT_NAME(Default_Object_Id) AS DF_CONSTRAINT_NAME         , "
            StrSql = StrSql & "       Rule_Object_Id                                               , "
            StrSql = StrSql & "       Is_Sparse                                                    , "
            StrSql = StrSql & "       Is_Column_Set                                                , "
            StrSql = StrSql & "       Is_Xml_Document                                              , "
            StrSql = StrSql & "       Xml_Collection_Id "
            StrSql = StrSql & "FROM   sysobjects "
            StrSql = StrSql & "       INNER JOIN syscolumns AS COLUMNS_1 "
            StrSql = StrSql & "        ON     COLUMNS_1.Id = sysobjects.id "
            StrSql = StrSql & "       INNER JOIN sys.columns AS COLUMNS_2 "
            StrSql = StrSql & "        ON     COLUMNS_2.object_id = COLUMNS_1.id "
            StrSql = StrSql & "        AND    COLUMNS_2.column_id = COLUMNS_1.colid "
            StrSql = StrSql & "       LEFT JOIN ( "
            StrSql = StrSql & "                    SELECT   object_id AS c_object_id, column_id AS c_column_id, "
            StrSql = StrSql & "                             [definition] AS c_definition, is_persisted AS c_is_persisted "
            StrSql = StrSql & "                    FROM  sys.computed_columns "
            StrSql = StrSql & "                 ) AS COMPUTED_FIELD "
            StrSql = StrSql & "        ON COMPUTED_FIELD.c_object_id = COLUMNS_1.id "
            StrSql = StrSql & "        AND COMPUTED_FIELD.c_column_id = COLUMNS_1.colid "
            StrSql = StrSql & "WHERE  sysobjects.XTYPE           = 'U'"

            StrSql = "(" & StrSql & ") AS VIEW_SCHEMA_ALL_COLUMNS"

            GetViewSchemaAllColumns = StrSql

        End Get
    End Property

    Private ReadOnly Property GetViewSchemaTriggers As String
        Get
            Dim StrSql As String

            StrSql = ""
            StrSql = StrSql & "SELECT SCHEMA_NAME(uid)          AS SCHEMA_NAME, "
            StrSql = StrSql & "       sysobjects.Name           AS FN_NAME    , "
            StrSql = StrSql & "       sysobjects.id             AS Object_Id  , "
            StrSql = StrSql & "       dbo.syscomments.ColId                   , "
            StrSql = StrSql & "       Text AS Script "
            StrSql = StrSql & "FROM   sysobjects "
            StrSql = StrSql & "       INNER JOIN dbo.syscomments "
            StrSql = StrSql & "       ON     dbo.syscomments.id = sysobjects.id "
            StrSql = StrSql & "WHERE  XTYPE                     = 'TR'"

            StrSql = "(" & StrSql & ") AS VIEW_SCHEMA_TRIGGERS"

            GetViewSchemaTriggers = StrSql
        End Get
    End Property

    Private ReadOnly Property GetViewSchemaFunctions As String
        Get
            Dim StrSql As String

            StrSql = ""
            StrSql = StrSql & "SELECT SCHEMA_NAME(uid)           AS SCHEMA_NAME, "
            StrSql = StrSql & "       sysobjects.Name            AS FN_NAME    , "
            StrSql = StrSql & "       sysobjects.id              AS Object_Id  , "
            StrSql = StrSql & "       dbo.syscomments.ColId                    , "
            StrSql = StrSql & "       Text AS Script "
            StrSql = StrSql & "FROM   sysobjects "
            StrSql = StrSql & "       INNER JOIN dbo.syscomments "
            StrSql = StrSql & "       ON     dbo.syscomments.id = sysobjects.id "
            StrSql = StrSql & "WHERE  XTYPE                     = 'FN'"

            StrSql = "(" & StrSql & ") AS VIEW_SCHEMA_FUNCTIONS"

            GetViewSchemaFunctions = StrSql
        End Get
    End Property

    Private ReadOnly Property GetViewSchemaViews As String
        Get
            Dim StrSql As String

            StrSql = ""
            StrSql = StrSql & "SELECT ALL_VIEWS.Object_Id, "
            StrSql = StrSql & "       ColId              , "
            StrSql = StrSql & "       Name                               AS VIEW_NAME  , "
            StrSql = StrSql & "       Text                               AS Script "
            StrSql = StrSql & "FROM   [sys].[all_views]                  AS ALL_VIEWS "
            StrSql = StrSql & "       INNER JOIN [sys].[all_sql_modules] AS SQL_MODULES "
            StrSql = StrSql & "       ON     SQL_MODULES.object_id = ALL_VIEWS.object_id "
            StrSql = StrSql & "       INNER JOIN dbo.syscomments "
            StrSql = StrSql & "       ON     dbo.syscomments.id = ALL_VIEWS.object_id "
            StrSql = StrSql & "WHERE  SCHEMA_NAME(schema_id)    = 'dbo'"

            StrSql = "(" & StrSql & ") AS VIEW_SCHEMA_VIEWS"

            GetViewSchemaViews = StrSql
        End Get
    End Property



    Private ReadOnly Property GetViewSchemaDatabases As String
        Get
            Dim StrSql As String

            StrSql = ""
            StrSql = StrSql & "SELECT * "
            StrSql = StrSql & "FROM   [sys].[databases]"

            StrSql = "(" & StrSql & ") AS VIEW_SCHEMA_DATABASES"

            GetViewSchemaDatabases = StrSql
        End Get
    End Property



    Private ReadOnly Property GetViewSchemaConstraintId As String
        Get
            Dim StrSql As String

            StrSql = ""
            StrSql = StrSql & "SELECT OBJECT_NAME(object_id)    AS TABLE_NAME , "
            StrSql = StrSql & "       NAME                      AS COLUMN_NAME, "
            StrSql = StrSql & "       TYPE_NAME(system_type_id) AS SYSTEM_TYPE, "
            StrSql = StrSql & "       TYPE_NAME(user_type_id)   AS USER_TYPE  , "
            StrSql = StrSql & "       SEED_VALUE                              , "
            StrSql = StrSql & "       INCREMENT_VALUE                         , "
            StrSql = StrSql & "       LAST_VALUE "
            StrSql = StrSql & "FROM   [sys].[identity_columns]"

            StrSql = "(" & StrSql & ") AS VIEW_CONSTRAINTS_ID"

            GetViewSchemaConstraintId = StrSql
        End Get
    End Property



    Private ReadOnly Property GetViewSchemaConstraintCk As String
        Get
            Dim StrSql As String

            StrSql = ""
            StrSql = StrSql & "SELECT OBJECT_NAME(COLUMNS.Object_Id)           AS TABLE_NAME     , "
            StrSql = StrSql & "       OBJECT_NAME(CHECK_CONSTRAINTS.Object_Id) AS CONSTRAINT_NAME, "
            StrSql = StrSql & "       COLUMNS.name                             AS COLUMN_NAME    , "
            StrSql = StrSql & "       CHECK_CONSTRAINTS.DEFINITION "
            StrSql = StrSql & "FROM   sys.check_constraints  AS CHECK_CONSTRAINTS "
            StrSql = StrSql & "       INNER JOIN sys.columns AS COLUMNS "
            StrSql = StrSql & "       ON     CHECK_CONSTRAINTS.parent_object_id = COLUMNS.object_id "
            StrSql = StrSql & "       AND    CHECK_CONSTRAINTS.parent_column_id = COLUMNS.column_id"

            StrSql = "(" & StrSql & ") AS VIEW_CONSTRAINTS_CK"

            GetViewSchemaConstraintCk = StrSql
        End Get
    End Property



    Private ReadOnly Property GetViewSchemaConstraintDf As String
        Get
            Dim StrSql As String

            StrSql = ""
            StrSql = StrSql & "SELECT OBJECT_NAME(COLUMNS.Object_Id)             AS TABLE_NAME     , "
            StrSql = StrSql & "       OBJECT_NAME(DEFAULT_CONSTRAINTS.Object_Id) AS CONSTRAINT_NAME, "
            StrSql = StrSql & "       COLUMNS.name                               AS COLUMN_NAME    , "
            StrSql = StrSql & "       [Definition]                                                 , "
            StrSql = StrSql & "       is_system_named         AS IsSystemNamed "
            StrSql = StrSql & "FROM   sys.default_constraints AS DEFAULT_CONSTRAINTS "
            StrSql = StrSql & "       INNER JOIN sys.columns  AS COLUMNS "
            StrSql = StrSql & "       ON     DEFAULT_CONSTRAINTS.parent_object_id = COLUMNS.object_id "
            StrSql = StrSql & "       AND    DEFAULT_CONSTRAINTS.parent_column_id = COLUMNS.column_id"

            StrSql = "(" & StrSql & ") AS VIEW_CONSTRAINTS_DF"

            GetViewSchemaConstraintDf = StrSql
        End Get
    End Property

    Private ReadOnly Property GetViewSchemaConstraintFk As String
        Get
            Dim StrSql As String

            StrSql = ""
            StrSql = StrSql & "SELECT FK.CONSTRAINT_NAME       AS Fk_Name       , "
            StrSql = StrSql & "       KU.TABLE_NAME            AS DependentTable, "
            StrSql = StrSql & "       KU.COLUMN_NAME           AS DependentCol  , "
            StrSql = StrSql & "       KU.ORDINAL_POSITION      AS DependentOrder, "
            StrSql = StrSql & "       KU2.TABLE_NAME           AS SourceTable   , "
            StrSql = StrSql & "       KU2.COLUMN_NAME          AS SourceCol     , "
            StrSql = StrSql & "       KU2.ORDINAL_POSITION     AS SourceOrder   , "
            StrSql = StrSql & "       Update_Rule, Delete_Rule, Match_Option    "
            StrSql = StrSql & "FROM   INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS FK "
            StrSql = StrSql & "       INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE KU "
            StrSql = StrSql & "       ON     KU.CONSTRAINT_NAME = FK.CONSTRAINT_NAME "
            StrSql = StrSql & "       INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE KU2 "
            StrSql = StrSql & "       ON     KU2.CONSTRAINT_NAME = FK.UNIQUE_CONSTRAINT_NAME "
            StrSql = StrSql & "       AND    KU.ORDINAL_POSITION = KU2.ORDINAL_POSITION"

            StrSql = "(" & StrSql & ") AS VIEW_CONSTRAINTS_FK"

            GetViewSchemaConstraintFk = StrSql
        End Get
    End Property

    Private ReadOnly Property GetViewSchemaConstraintPk As String
        Get
            Dim StrSql As String

            StrSql = ""
            StrSql = StrSql & "SELECT Table_Schema             AS [Schema]  , "
            StrSql = StrSql & "       Constraint_Catalog       AS DbName    , "
            StrSql = StrSql & "       Table_Name               AS TableName , "
            StrSql = StrSql & "       constraint_name          AS PkName    , "
            StrSql = StrSql & "       Column_Name              AS ColumnName, "
            StrSql = StrSql & "       Ordinal_Position         AS ColumnOrder "
            StrSql = StrSql & "FROM   INFORMATION_SCHEMA.KEY_COLUMN_USAGE "
            StrSql = StrSql & "WHERE  OBJECTPROPERTY(OBJECT_ID(constraint_name), 'IsPrimaryKey') = 1 "

            StrSql = "(" & StrSql & ") AS VIEW_CONSTRAINTS_PK"

            GetViewSchemaConstraintPk = StrSql
        End Get
    End Property


    Private ReadOnly Property GetViewSchemaConstraintUq As String
        Get
            Dim StrSql As String

            StrSql = ""
            StrSql = StrSql & "SELECT OBJECT_NAME(kc.parent_object_id) AS TABLE_NAME, "
            StrSql = StrSql & "       kc.name                          AS PK_NAME   , "
            StrSql = StrSql & "       c.NAME                           AS COLUMN_NAME "
            StrSql = StrSql & "FROM   sys.key_constraints kc "
            StrSql = StrSql & "       INNER JOIN sys.index_columns ic "
            StrSql = StrSql & "       ON     kc.parent_object_id = ic.object_id "
            StrSql = StrSql & "       INNER JOIN sys.columns c "
            StrSql = StrSql & "       ON     ic.object_id = c.object_id "
            StrSql = StrSql & "       AND    ic.column_id = c.column_id "
            StrSql = StrSql & "WHERE  kc.type             = 'UQ'"
            StrSql = StrSql & "       AND OBJECT_NAME(kc.parent_object_id) NOT IN ( "
            StrSql = StrSql & "                          'VIEW_CONSTRAINTS_UQ', 'VIEW_CONSTRAINTS_PK', 'VIEW_CONSTRAINTS_FK', "
            StrSql = StrSql & "                          'VIEW_CONSTRAINTS_DF', 'VIEW_CONSTRAINTS_CK', 'VIEW_CONSTRAINTS_ID', 'VIEW_SCHEMA_DATABASES', "
            StrSql = StrSql & "                          'VIEW_SCHEMA_VIEWS', 'VIEW_SCHEMA_FUNCTIONS', 'VIEW_SCHEMA_TRIGGERS', 'VIEW_SCHEMA_ALL_COLUMNS' "
            StrSql = StrSql & "                       )"

            StrSql = "(" & StrSql & ") AS VIEW_CONSTRAINTS_UQ"

            GetViewSchemaConstraintUq = StrSql
        End Get
    End Property


    Public Function GenerateCommandList(ByVal FromDataBase As String, _
                                        Optional ByVal ToDataBase As String = "", _
                                        Optional SelectCommands As List(Of CommandType) = Nothing) As List(Of String)

        'Dim MapScriptGroup As New MapList
        Dim CommandList As New List(Of String)
        Dim SavedDatabaseName As String

        'Salvando nome do BD:
        SavedDatabaseName = Me.Connection.GetCurrentConnection.Database
        Connection.GetCurrentConnection.ChangeDatabase(FromDataBase)

        If ToDataBase.IsEmpty Then
            ToDataBase = FromDataBase
        End If

        RaiseEvent LogEvent(Me, New LogEventArgs("GenerateCommandList", "Generating command list"))

        CommandList.Add("USE " & "master".Squared)

        If SelectCommands.Any(Function(x) x = CommandType.Database) Or SelectCommands Is Nothing Then
            CommandList.Add(Me.MakeDatabase(FromDataBase, ToDataBase))
            CommandList.Add("USE " & ToDataBase.Squared)
        End If

        If SelectCommands.Any(Function(x) x = CommandType.Table) Or SelectCommands Is Nothing Then _
            CommandList.Add(Me.MakeTables)

        If SelectCommands.Any(Function(x) x = CommandType.DefaultConstraint) Or SelectCommands Is Nothing Then _
            CommandList.Add(Me.MakeDefaultConstraints)

        If SelectCommands.Any(Function(x) x = CommandType.CheckConstraint) Or SelectCommands Is Nothing Then _
            CommandList.Add(Me.MakeCheckConstraints)

        If SelectCommands.Any(Function(x) x = CommandType.UniqueConstraint) Or SelectCommands Is Nothing Then _
            CommandList.Add(Me.MakeUniqueConstraints)

        If SelectCommands.Any(Function(x) x = CommandType.PrimaryKey) Or SelectCommands Is Nothing Then _
            CommandList.Add(Me.MakePrimaryKeys)

        If SelectCommands.Any(Function(x) x = CommandType.ForeignKey) Or SelectCommands Is Nothing Then _
            CommandList.Add(Me.MakeForeignKeys)

        If SelectCommands.Any(Function(x) x = CommandType.StoredProcedure) Or SelectCommands Is Nothing Then _
            CommandList.Add(Me.MakeFunctions)

        If SelectCommands.Any(Function(x) x = CommandType.Trigger) Or SelectCommands Is Nothing Then _
            CommandList.Add(Me.MakeTriggers)

        If SelectCommands.Any(Function(x) x = CommandType.View) Or SelectCommands Is Nothing Then _
            CommandList.Add(Me.MakeViews)


        Connection.GetCurrentConnection.ChangeDatabase(SavedDatabaseName) 'Recuperando nome do BD:


        Return CommandList

    End Function


    Public Function GenerateFormatedScript(ByVal FromDataBase As String, _
                                           Optional ByVal ToDataBase As String = "") As String

        Dim Scripts As String
        Dim CommandList As New List(Of String)

        RaiseEvent LogEvent(Me, New LogEventArgs("GenerateFormatedScript", "Generating commands script"))

        CommandList = GenerateCommandList(FromDataBase, ToDataBase)

        Scripts = CommandList.JoinWith(vbNewLine & "GO" & vbNewLine & vbNewLine) & _
                    vbNewLine & "GO" & vbNewLine & vbNewLine

        Return Scripts
    End Function

    Private Shadows Sub VerificarSeUsuarioCancelou()
        If Not Me.BackgroundWorker.IsNull AndAlso Me.BackgroundWorker.CancellationPending Then
            Throw New Exception("Procedimento cancelado pelo usuário")
        End If
    End Sub

End Class


