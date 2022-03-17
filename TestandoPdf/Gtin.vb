Imports System.Text.RegularExpressions


Public Class Gtin

    Public Enum TipoCodigoBarras
        NaoDefinido
        EAN_8
        EAN_13
        GTIN_8
        GTIN_12
        GTIN_13
        GTIN_14
    End Enum

    ''' <summary>
    ''' Valida códigos de barras padrões GTIN/EAN
    ''' </summary>
    ''' <param name="CodigoBarras"></param>
    ''' <param name="TipoAValidar"></param>
    ''' <returns></returns>
    Public Shared Function ValidarCodigoDeBarras(CodigoBarras As String, Optional TipoAValidar As TipoCodigoBarras = TipoCodigoBarras.NaoDefinido) As Boolean

        Dim vCodigoBarras As String = CodigoBarras.ToString.Trim


        If vCodigoBarras.IsEmpty Then _
            Return False


        Dim AuxValorInt As Integer = Integer.MinValue  ' <-- Variável auxiliar utilizada apenas para validar o TryParse

        Select Case vCodigoBarras.Length

            Case 8
                If TipoAValidar <> TipoCodigoBarras.EAN_8 And TipoAValidar <> TipoCodigoBarras.NaoDefinido Then _
                    Return False


                ' <-- Verifica se há somente números no código
                If Integer.TryParse(vCodigoBarras.Substring(0, 7), AuxValorInt) Then _
                    If vCodigoBarras(vCodigoBarras.Length - 1).ToString = CalcularDigitoEAN8(vCodigoBarras.Substring(0, 7)) Then _
                        Return True

                Exit Select

            Case 13

                If TipoAValidar <> TipoCodigoBarras.EAN_13 And TipoAValidar <> TipoCodigoBarras.NaoDefinido Then _
                    Return False

                '// <-- Verifica se há somente números no código
                If Integer.TryParse(vCodigoBarras.Substring(0, 12), AuxValorInt) Then _
                    If vCodigoBarras(vCodigoBarras.Length - 1).ToString = CalcularDigitoEAN13(vCodigoBarras.Substring(0, 12)) Then _
                        Return True
                Exit Select

        End Select





        If Not TipoAValidar.ToString().ToUpper().Contains("GTIN") And TipoAValidar <> TipoCodigoBarras.NaoDefinido Then _
                Return False

        If ValidaCodigoGTIN(CodigoBarras) Then _
            Return True



        Return False

    End Function


    Public Shared Function CalcularDigitoEAN13(codigo As String) As String

        Dim nPeso As Integer = 3
        Dim nSoma As Double = 0
        Dim nMaior As Double
        Dim nDigito As Integer
        Dim result As String = String.Empty

        For nX As Integer = 11 To 0 Step -1
            nSoma = nSoma + Integer.Parse(codigo(nX).ToString) * nPeso

            If nPeso = 3 Then

                nPeso = 1
            Else
                nPeso = 3
            End If
        Next

        nMaior = ((Math.Truncate(nSoma / 10) + 1) * 10)
        nDigito = Convert.ToInt32(Math.Truncate(nMaior) - Math.Truncate(nSoma))

        If nDigito = 10 Then _
            nDigito = 0

        result = nDigito.ToString

        Return result

    End Function



    Public Shared Function CalcularDigitoEAN8(codigo As String) As String

        Dim nPeso As Integer = 3
        Dim nSoma As Double = 0
        Dim nMaior As Double
        Dim nDigito As Integer
        Dim result As String = String.Empty

        For nX As Integer = 6 To 0 Step -1
            nSoma = nSoma + Integer.Parse(codigo(nX).ToString()) * nPeso

            If nPeso = 3 Then
                nPeso = 1
            Else
                nPeso = 3
            End If
        Next

        nMaior = ((Math.Truncate(nSoma / 10) + 1) * 10)
        nDigito = Convert.ToInt32(Math.Truncate(nMaior) - Math.Truncate(nSoma))

        If nDigito = 10 Then _
            nDigito = 0

        result = nDigito.ToString

        Return result

    End Function


    ''' <summary>
    ''' Validações de Códigos de Barra padrão GTIN, 8, 12, 13 e 14
    ''' </summary>
    ''' <param name="codigoGTIN">Cógigo GTIN 8,12,13,14</param>
    ''' <returns>True se válido</returns>
    Private Shared Function ValidaCodigoGTIN(codigoGTIN As String) As Boolean

        If codigoGTIN <> (New Regex("[^0-9]")).Replace(codigoGTIN, String.Empty) Then _
                Return False

        Select Case codigoGTIN.Length
            Case 8
                codigoGTIN = "000000" + codigoGTIN
                Exit Select
            Case 12
                codigoGTIN = "00" + codigoGTIN
                Exit Select
            Case 13
                codigoGTIN = "0" + codigoGTIN
                Exit Select
            Case 14
                Exit Select
            Case Else
                Return False
        End Select


        'Calculando dígito verificador
        Dim a(13) As Integer
        a(0) = Integer.Parse(codigoGTIN(0).ToString) * 3
        a(1) = Integer.Parse(codigoGTIN(1).ToString)
        a(2) = Integer.Parse(codigoGTIN(2).ToString) * 3
        a(3) = Integer.Parse(codigoGTIN(3).ToString)
        a(4) = Integer.Parse(codigoGTIN(4).ToString) * 3
        a(5) = Integer.Parse(codigoGTIN(5).ToString)
        a(6) = Integer.Parse(codigoGTIN(6).ToString) * 3
        a(7) = Integer.Parse(codigoGTIN(7).ToString)
        a(8) = Integer.Parse(codigoGTIN(8).ToString) * 3
        a(9) = Integer.Parse(codigoGTIN(9).ToString)
        a(10) = Integer.Parse(codigoGTIN(10).ToString) * 3
        a(11) = Integer.Parse(codigoGTIN(11).ToString)
        a(12) = Integer.Parse(codigoGTIN(12).ToString) * 3

        Dim sum As Integer = a(0) + a(1) + a(2) + a(3) + a(4) + a(5) + a(6) + a(7) + a(8) + a(9) + a(10) + a(11) + a(12)
        Dim check As Integer = (10 - (sum Mod 10)) Mod 10

        'Checando
        Dim last As Integer = Integer.Parse(codigoGTIN(13).ToString)
        Return (check = last)

    End Function


End Class










        

        


        