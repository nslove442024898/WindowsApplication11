Module Module1
    Public Function getRndN(ByVal l As Double, ByVal u As Double) As Double
        'l为下限，u为上限
        Randomize()

        Return Int(Rnd() * (u - l) + l) + Int(Rnd() + 0.5) * 0.5
    End Function
    Function LowLmt(ByVal lb As Double, ByVal ub As Double, ByVal cv As Double) As Double
        Return getPosN(lb - (ub - lb) * cv)
    End Function
    Function UpLmt(ByVal lb As Double, ByVal ub As Double, ByVal cv As Double) As Double
        Return ub + (ub - lb) * cv
    End Function
    Function getRndN(ByVal l As Double, ByVal u As Double, ByVal lastN As Double, ByVal cv As Double) As Double

        Randomize()
        Do While True
            Dim x = Int(Rnd() * (u - l) + l) + Int(Rnd() + 0.5) * 0.5
            If System.Math.Abs(x - lastN) < 5 Then Return x


        Loop

    End Function

    Function getRndInt(ByVal l As Double, ByVal u As Double) As Integer
        Randomize()

        Return Int(Rnd() * (u - l) + l)
    End Function
    Function getPosN(ByVal n As Double) As Double
        If n > 0 Then
            Return n
        Else
            Return 0
        End If
    End Function
End Module
