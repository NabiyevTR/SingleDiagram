
Module MyMod


    'округление
    ' Если больше единцы, то округляем до n знаков после запятой, если меньше то огругляем до n-значащих цифр
    Function RoundSign(ByRef x As Double, ByRef n As Integer) As Double


        Dim k As Integer = 1

        If x < 0 Then k = -1

        Dim xabs As Double = Math.Abs(x)

        If xabs = 0 Then
            Return (x)
            Exit Function
        End If

        If xabs >= 1 Then
            Return (Math.Round(x, n))
        Else
            Dim digits As Integer = n - RoundUp0(Math.Log10(xabs)) - 1
            Return (k * Math.Round(xabs * Math.Pow(10, digits)) / Math.Pow(10, digits))
        End If



    End Function

    Function RoundUp0(ByRef x As Double) As Double
        Dim k As Integer = 1
        If x < 0 Then k = -1

        Dim xabs As Double = Math.Abs(x)

        If xabs - Math.Round(xabs, 0) > 0 Then
            Return (k * RoundSign(xabs + 1, 0))
        Else : Return (k * RoundSign(xabs, 0))
        End If



    End Function




End Module
