Imports System.Runtime.CompilerServices
Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Reflection

Imports System.Xml
Imports System.Xml.Linq
Imports System.Xml.Schema
Imports System.Xml.Serialization

Module Utils

    '*************** EXTENSION METHODS ***************


    ''' <summary>
    ''' Verifica se valor é nulo ou vazio.
    ''' </summary>
    ''' <returns>Retorna verdadeiro se valor for Null (Nothing) ou Empty.</returns>
    <Extension()>
    Public Function IsEmpty(ByVal SelfObj As String) As Boolean

        Return String.IsNullOrEmpty(SelfObj)
    End Function

    ''' <summary>
    ''' Retorna um valor DEFAULT caso o valor da string seja EMPTY.
    ''' </summary>
    ''' <param name="SelfObj"></param>
    ''' <param name="DefaultValue">Valor padrão a ser retornado</param>
    ''' <returns>Retorna uma string</returns>
    <Extension()>
    Public Function IfIsEmpty(ByVal SelfObj As String, ByVal DefaultValue As String) As String

        Return IIf(String.IsNullOrEmpty(SelfObj), DefaultValue, SelfObj)
    End Function

    ''' <summary>
    ''' Retorna um valor DEFAULT caso o valor do objeto seja NULL.
    ''' </summary>
    ''' <param name="SelfObj"></param>
    ''' <param name="DefaultValue">Valor padrão a ser retornado</param>
    ''' <returns>Retorna um objeto</returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function IfIsNull(ByVal SelfObj As Object, ByVal DefaultValue As Object) As Object

        Return IIf(SelfObj Is Nothing, DefaultValue, SelfObj)
    End Function

    ''' <summary>
    ''' Obtém uma quantidade de caracteres à esquerda.
    ''' </summary>
    ''' <param name="IntCorte">Quantidade de caracteres.</param>
    ''' <returns>Retorna sequência de caracteres indicada.</returns>
    <Extension()>
    Public Function TakeLeft(ByVal SelfObj As String, ByVal IntCorte As Integer) As String

        Return Microsoft.VisualBasic.Left(SelfObj, IntCorte)
    End Function

    ''' <summary>
    ''' Obtém uma quantidade de caracteres à direita.
    ''' </summary>
    ''' <param name="IntCorte">Quantidade de caracteres.</param>
    ''' <returns>Retorna sequência de caracteres indicada.</returns>
    <Extension()>
    Public Function TakeRight(ByVal SelfObj As String, ByVal IntCorte As Integer) As String
        Return Microsoft.VisualBasic.Right(SelfObj, IntCorte)
    End Function

    <Extension()>
    Public Function FormatTo(ByVal SelfObj As String, ByVal ParamArray args() As Object) As String
        Return String.Format(SelfObj, args)
    End Function

    <Extension()>
    Public Function FormatTo(ByVal SelfObj As String, ByVal provider As System.IFormatProvider, ByVal ParamArray args() As Object) As String
        Return String.Format(provider, SelfObj, args)
    End Function

    ''' <summary>
    ''' Adiciona aspas simples à string.
    ''' </summary>
    ''' <param name="SelfObj"></param>
    ''' <returns>Retorna a string com aspas simples</returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function Quote(ByVal SelfObj As String) As String
        Return Chr(39) & SelfObj & Chr(39)
    End Function

    <Extension()>
    Public Function Bracket(ByVal SelfObj As String) As String
        Return Chr(40) & SelfObj & Chr(41)
    End Function

    <Extension()>
    Public Function Squared(ByVal SelfObj As String) As String
        Return Chr(91) & SelfObj & Chr(93)
    End Function


    ''' <summary>
    ''' Converte o XDocument atual em um XmlDocument.
    ''' </summary>
    ''' <returns>Retorna um objeto do tipo XmlDocument.</returns>
    <Extension()>
    Public Function ToXmlDocument(ByVal SelfObj As XDocument) As XmlDocument

        Dim xmlDoc As New XmlDocument

        xmlDoc.Load(SelfObj.CreateReader)

        Return xmlDoc
    End Function

    ''' <summary>
    ''' Converte o XmlDocument atual em um XDocument.
    ''' </summary>
    ''' <returns>Retorna um objeto do tipo XDocument.</returns>
    <Extension()>
    Public Function ToXDocument(ByVal SelfObj As XmlDocument) As XDocument

        Return XDocument.Parse(SelfObj.OuterXml, LoadOptions.PreserveWhitespace)
    End Function



    ''' <summary>
    ''' Converte o XElement atual em um XmlElement.
    ''' </summary>
    ''' <returns>Retorna um objeto do tipo XmlElement.</returns>
    <Extension()>
    Public Function ToXmlElement(ByVal SelfObj As XElement) As XmlElement

        Dim xmlDoc As New XmlDocument
        xmlDoc.PreserveWhitespace = True
        Return xmlDoc.ReadNode(SelfObj.CreateReader())
    End Function

    ''' <summary>
    ''' Converte o XmlElement atual em um XElement.
    ''' </summary>
    ''' <returns>Retorna um objeto do tipo XElement.</returns>
    <Extension()>
    Public Function ToXElement(ByVal SelfObj As XmlElement) As XElement

        Return XElement.Parse(SelfObj.OuterXml, LoadOptions.PreserveWhitespace)
    End Function


    ''' <summary>
    ''' O mesmo que String.Join(separator, array()). Funciona apenas com List(of String).
    ''' </summary>
    ''' <param name="SelfObj"></param>
    ''' <param name="separator"></param>
    ''' <returns>Retorna uma string contendo o conteúdo de cada elemento do array especificado, separados pelo argumento "separator".</returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function JoinWith(ByVal SelfObj As List(Of String), ByVal separator As String) As String

        Return String.Join(separator, SelfObj.ToArray)
    End Function

    ''' <summary>
    ''' Transforma uma lista de inteiros em uma lista de strings, separados pelo argumento [separator].
    ''' </summary>
    ''' <param name="SelfObj"></param>
    ''' <param name="separator">Valor inserido entre os itens da lista</param>
    ''' <returns>Retorna uma string contendo o conteúdo de cada elemento do array especificado, separados pelo argumento "separator".</returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function JoinWith(ByVal SelfObj As List(Of Integer), ByVal Separator As String) As String

        Dim Lista As New List(Of String)

        SelfObj.ForEach(Sub(x)
                            Lista.Add(x)
                        End Sub)

        Return String.Join(Separator, Lista.ToArray)
    End Function

    <Extension()>
    Public Function Add(ByVal SelfObj As List(Of String), ByVal anotherList As List(Of String)) As List(Of String)

        anotherList.ForEach(Sub(item)
                                SelfObj.Add(item)
                            End Sub)

        Return SelfObj
    End Function

    <Extension()>
    Public Sub ForEach(Of T1, T2)(ByRef SelfObj As Dictionary(Of T1, T2), ByRef Action As System.Action(Of T1, T2))
        For Each xItem As KeyValuePair(Of T1, T2) In SelfObj
            Action(xItem.Key, xItem.Value)
        Next
    End Sub


    ''' <summary>
    ''' O mesmo que String.Join(separator, array()).
    ''' </summary>
    ''' <param name="SelfObj"></param>
    ''' <param name="separator"></param>
    ''' <returns>Retorna uma string contendo o conteúdo de cada elemento do array especificado, separados pelo argumento "separator".</returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function JoinWith(ByVal SelfObj As String(), ByVal separator As String) As String

        Return String.Join(separator, SelfObj)
    End Function


    ''' <summary>
    ''' Retorna a parte chave (nome) de um "Enum" instanciado.
    ''' </summary>
    ''' <param name="SelfObj"></param>
    ''' <returns>Retorna nome/chave da instância de Enum</returns>
    <Extension()>
    Public Function GetName(ByVal SelfObj As System.Enum) As String

        Return [Enum].GetName(SelfObj.GetType, SelfObj)
    End Function


    <Extension()>
    Public Function IsNull(ByVal SelfObj As Object) As Boolean

        Return (SelfObj Is DBNull.Value Or SelfObj Is Nothing)
    End Function


    

    ''' <summary>
    ''' Verifica se existe a string dentro de Collection.
    ''' </summary>
    <Extension()>
    Public Function ExistsIn(ByRef SelfObj As String, _
                               ByRef Collection As IEnumerable) As Boolean

        'IEnumerable
        For Each Item As String In Collection
            If Item = SelfObj Then
                Return True
            End If
        Next

        Return False
    End Function

    ''' <summary>
    ''' Verifica se existe a string dentro de [ParamArray].
    ''' </summary>
    <Extension()>
    Public Function ExistsIn(ByRef SelfObj As String, _
                               ParamArray Args As String()) As Boolean



        'IEnumerable
        For Each Item As String In Args.ToList
            If Item = SelfObj Then
                Return True
            End If
        Next

        Return False
    End Function


    '************************* EXTENSION METHODS *******************************



    Public Function ConectarServidor(StrNomeServidor As String, StrNomeBaseDados As String) As ConnectionManager
        Dim CommandsList As System.Collections.Generic.List(Of String)
        Dim StrConexao As String


        CommandsList = New System.Collections.Generic.List(Of String)
        CommandsList.Add("SET LANGUAGE 'Português (Brasil)'")
        CommandsList.Add("SET LOCK_TIMEOUT 5000")

        StrNomeBaseDados = StrNomeBaseDados.Trim
        StrNomeServidor = StrNomeServidor.Trim

        StrConexao = "Initial Catalog=" & StrNomeBaseDados & ";" & _
                     "Data Source=" & StrNomeServidor & ";" & _
                     "User ID=preview;" & _
                     "Password=919985;" & _
                     "Connect Timeout=0;" & _
                     "Application Name='CloneMaker'"

        'StrConexao = "Initial Catalog=" & StrNomeBaseDados & ";" & _
        '             "Data Source=" & StrNomeServidor & ";" & _
        '             "User ID=preview;" & _
        '             "Password=admin@atende1317;" & _
        '             "Connect Timeout=0;" & _
        '             "Application Name='Teste'"

        ConnectionManager.CreateInstance(StrConexao, CommandsList)
        Return ConnectionManager.GetInstance
    End Function



End Module
