Public Structure struDgv波速表设置
    Public tcbh As String
    Public tcmc As String
    Public loLmt横波 As Double
    Public UpLmt横波 As Double
    Public cv横波 As Double
    Public loLmt纵波 As Double
    Public UpLmt纵波 As Double
    Public cv纵波 As Double
    Sub New(intcbh As String, intcmc As String, inhengbololmt As Double, inhengbouplmt As Double, inhengbocv As Double, inzonglolmt As Double,
            inzonguplmt As Double, inzongcv As Double, ByRef success As Boolean)
        Try
            With Me
                .tcbh = intcbh
                .tcmc = intcmc
                .loLmt横波 = inhengbololmt
                .UpLmt横波 = inhengbouplmt
                .cv横波 = inhengbocv

                .loLmt纵波 = inzonglolmt
                .UpLmt纵波 = inzonguplmt
                .cv纵波 = inzongcv
            End With
        Catch ex As Exception
            MsgBox("读入波速表设置出错" & Err.Description)
            success = False
        End Try
    End Sub
End Structure
Public Structure struDgv取芯和RQD
    Public tcbh As String
    Public tcmc As String
    Public qxllo As Double
    Public RQDlo As Double
    Public qxlup As Double
    Public RQDup As Double
    Sub New(intcbh As String, intcmc As String, inqxllo As Double, inqxlup As Double, inrqdlo As Double, inrqdup As Double)
        Try
            With Me
                .tcbh = intcbh
                .tcmc = intcmc
                .qxllo = inqxllo
                .qxlup = inqxlup
                .RQDlo = inrqdlo
                .RQDup = inrqdup
            End With
        Catch ex As Exception

        End Try
    End Sub
End Structure
Public Structure 钻孔编号高程数据结构
    Dim zkbh As String
    Dim zkbg As Double
    Dim zkx As Double
    Dim zky As Double
    Dim zksd As Double
    Dim ZKKSRQ As String '钻孔开始日期
    Dim ZKZZRQ As String '钻孔终止日期
    Dim ZKZJLX As String '钻机类型

End Structure
Public Structure 地层数据结构
    '一般参数
    Dim zkbh As String
    Dim zkbg As Double
    Dim TCZCBH As Short
    Dim TCYCBH As Short
    Dim tccycbh As Short
    Dim tcmc As String
    Dim tclm As String

    Dim TCYS As String ' ADDED BY NANS @20201108
    Dim TCKSX As String 'ADDED BY NANS @20201108
    Dim TCSID As String 'ADDED BY NANS @20201108
    Dim TCMS As String 'ADDED BY NANS @20201108
    Dim TCMSD As String 'ADDED BY NANS @20201108

    Dim cengdingbiaogao As Double
    Dim cengdibiaogao As Double
    Dim tchd As Double
    Dim cengdingshendu As Double '土层层顶深度
    Dim cengdishengdu As Double '土层层底深度

    Sub New(v As Double)
        zkbh = ""
        zkbg = 0
        TCZCBH = 0
        TCYCBH = 0
        tccycbh = 0
        tcmc = ""
        tclm = ""
        '地层空间分布参数、


        cengdingbiaogao = 0
        cengdibiaogao = 0
        tchd = 0
        cengdingshendu = 0
        cengdishengdu = 0
    End Sub
End Structure
Public Structure struBosu
    Public dsd As Double
    Public bosuh As Double
    Public bosuz As Double
    Sub New(indsd As Double, inbosuh As Double, inbosuz As Double)
        Try
            With Me
                .dsd = indsd
                .bosuh = inbosuh
                .bosuz = inbosuz
            End With
        Catch ex As Exception
            MsgBox("初始化波速出错")
        End Try
    End Sub
End Structure
Public Structure stru取芯RQD
    Public dsd As Double
    Public qxl As Double
    Public rqd As Double
    Sub New(indsd As Double, inqxl As Double, inrqd As Double)
        Try
            With Me
                .dsd = indsd
                .qxl = inqxl
                .rqd = inrqd
            End With
        Catch ex As Exception

        End Try

    End Sub
End Structure
Public Class qyywcsOptor
    Enum eDrection
        increase = 1
        decrease = 2
        none = 3
    End Enum
    Dim UnitLen As Double
    Dim tranLen As Integer



    Function getThisDr(ByVal lastDr As eDrection) As eDrection

        Select Case lastDr
            Case eDrection.decrease
                Return eDrection.increase
            Case eDrection.increase
                Return eDrection.decrease
            Case eDrection.none
                Return getRndIntNamong(1, 2)
        End Select
    End Function
    Function getRndIntNamong(ByVal n1 As Integer, ByVal n2 As Integer) As Integer
        If n1 > 3 Or n1 < 1 Then Return n1
        If n2 > 3 Or n2 < 1 Then Return n2

        Do While True
            Dim n As Integer = getRndIntN(1, 3)
            If n = n1 Or n = n2 Then
                Return n
                Exit Do
            End If
        Loop
    End Function
    Function getRndIntN(ByVal l As Integer, ByVal u As Integer) As Integer
        Randomize()
        Return Int(Rnd() * (u - l) + l)
    End Function

    Structure TypeDt
        Dim dsd As Double
        Dim dtgc As Double
        Dim dtjs As Double
        Dim dtxzjs As Double
        Dim dtlx As eDtlx
        Dim dtcd As Double
    End Structure
    Structure typetc
        Dim tcbh As String
        Dim tclen As Double
        Dim tcdingSD As Double
        Dim gcsy As Integer
        Dim tcdsd As Double
        Dim tcType As QyETcType
    End Structure
    Enum eDtlx
        未定义 = 0
        N10 = 1
        N635 = 2
        N120 = 3
    End Enum

    Structure typeDgvTbl
        Dim tcbh As String
        Dim tcmc As String
        Dim cslx As eDtlx
        Dim loLmt As Double
        Dim UpLmt As Double
        Dim cv As Double
        Dim scrubInterval As Integer
        Dim scrubCount As Integer
        Dim addCount As Integer
        Dim specifiedZks As String
        Dim MinLenght As Double
        Dim testCount As Integer
        Dim qycd As Double
        Dim qyCount As Integer
    End Structure
    Structure relatedZK
        Dim zkbh As String
    End Structure
    Structure TypeBG
        Dim dsd As Double
        Dim bgjs As Double
        Dim bgxzjs As Double
        Dim bggc As Double

    End Structure
    Structure TypeQuyang
        Dim qybh As String
        Dim qydsd As Double
        Dim qycd As Double

    End Structure
    Structure TypeShuiwei
        Dim sch As Integer
        Dim swsd As Double
        Dim slx As Double
    End Structure
    Enum QyETcType
        soil = 1
        rock = 2
    End Enum
    Sub getZKQyArr(ByRef zkqyArr() As TypeQuyang, ByVal tcTbl() As typetc, ByRef dtbl() As typeDgvTbl, ByVal qyMinsd As Double, ByVal qySpace As Double, ByVal tc As QyETcType)


        For i As Integer = 0 To tcTbl.Length - 1
            If tcTbl(i).tclen > dtbl(getIndexInDgvTbl(dtbl, tcTbl(i).tcbh)).MinLenght And dtbl(getIndexInDgvTbl(dtbl, tcTbl(i).tcbh)).qyCount > 0 And tcTbl(i).tcType = tc Then
                Dim m As Integer = 1
                Do While True
                    Select Case m
                        Case 1
                            If tcTbl(i).tcdingSD + 1.2 < tcTbl(i).tcdsd And tcTbl(i).tcdingSD + 0.6 > qyMinsd And dtbl(getIndexInDgvTbl(dtbl, tcTbl(i).tcbh)).qyCount > 0 Then
                                If Not (zkqyArr.Length = 1 And zkqyArr(0).qydsd = 0) Then ReDim Preserve zkqyArr(zkqyArr.Length)
                                zkqyArr(zkqyArr.Length - 1).qydsd = tcTbl(i).tcdingSD + 0.6
                                If zkqyArr.Length = 1 Then
                                    zkqyArr(0).qybh = 1
                                Else
                                    zkqyArr(zkqyArr.Length - 1).qybh = zkqyArr(zkqyArr.Length - 2).qybh + 1
                                End If

                                If dtbl(getIndexInDgvTbl(dtbl, tcTbl(i).tcbh)).qycd > 0 Then
                                    zkqyArr(zkqyArr.Length - 1).qycd = dtbl(getIndexInDgvTbl(dtbl, tcTbl(i).tcbh)).qycd
                                Else
                                    zkqyArr(zkqyArr.Length - 1).qycd = 0.2
                                End If

                                m = m + 1
                                dtbl(getIndexInDgvTbl(dtbl, tcTbl(i).tcbh)).qyCount = dtbl(getIndexInDgvTbl(dtbl, tcTbl(i).tcbh)).qyCount - 1
                            Else
                                GoTo 100
                            End If
                        Case Else
                            If zkqyArr(zkqyArr.Length - 1).qydsd + qySpace + dtbl(getIndexInDgvTbl(dtbl, tcTbl(i).tcbh)).qycd <= tcTbl(i).tcdsd And dtbl(getIndexInDgvTbl(dtbl, tcTbl(i).tcbh)).qyCount > 0 Then
                                If Not (zkqyArr.Length = 1 And zkqyArr(0).qydsd = 0) Then ReDim Preserve zkqyArr(zkqyArr.Length)
                                zkqyArr(zkqyArr.Length - 1).qydsd = zkqyArr(zkqyArr.Length - 2).qydsd + qySpace
                                zkqyArr(zkqyArr.Length - 1).qybh = zkqyArr(zkqyArr.Length - 2).qybh + 1
                                If dtbl(getIndexInDgvTbl(dtbl, tcTbl(i).tcbh)).qycd > 0 Then
                                    zkqyArr(zkqyArr.Length - 1).qycd = dtbl(getIndexInDgvTbl(dtbl, tcTbl(i).tcbh)).qycd
                                Else
                                    zkqyArr(zkqyArr.Length - 1).qycd = 0.2
                                End If
                                m = m + 1
                                dtbl(getIndexInDgvTbl(dtbl, tcTbl(i).tcbh)).qyCount = dtbl(getIndexInDgvTbl(dtbl, tcTbl(i).tcbh)).qyCount - 1
                            Else
                                Exit Do
                            End If
                    End Select
                Loop
            End If
100:    Next
    End Sub
    Sub getZkBgArr(ByRef zkBgArr() As TypeBG, ByVal tcTbl() As typetc, ByRef dtbl() As typeDgvTbl, ByVal bgMinsd As Double, ByVal bgspace As Double)
        For i As Integer = 0 To tcTbl.Length - 1
            If dtbl(getIndexInDgvTbl(dtbl, tcTbl(i).tcbh)).UpLmt > 0 And tcTbl(i).tclen > dtbl(getIndexInDgvTbl(dtbl, tcTbl(i).tcbh)).MinLenght And dtbl(getIndexInDgvTbl(dtbl, tcTbl(i).tcbh)).testCount > 0 Then
                Dim m As Integer = 1
                Do While True
                    'If Not (zkBgArr.Length = 1 And zkBgArr(0).bggc = 0) Then ReDim Preserve zkBgArr(zkBgArr.Length)
                    Select Case m
                        Case 1
                            If tcTbl(i).tcdingSD + 0.6 < tcTbl(i).tcdsd And tcTbl(i).tcdingSD + 0.2 >= bgMinsd Then
                                If Not (zkBgArr.Length = 1 And zkBgArr(0).bggc = 0) Then ReDim Preserve zkBgArr(zkBgArr.Length)
                                zkBgArr(zkBgArr.Length - 1).dsd = tcTbl(i).tcdingSD + 0.2
                                zkBgArr(zkBgArr.Length - 1).bggc = zkBgArr(zkBgArr.Length - 1).dsd + 0.8
                                zkBgArr(zkBgArr.Length - 1).bgjs = getRndInt(dtbl(getIndexInDgvTbl(dtbl, tcTbl(i).tcbh)).loLmt, dtbl(getIndexInDgvTbl(dtbl, tcTbl(i).tcbh)).UpLmt)
                                m = m + 1
                                dtbl(getIndexInDgvTbl(dtbl, tcTbl(i).tcbh)).testCount = dtbl(getIndexInDgvTbl(dtbl, tcTbl(i).tcbh)).testCount - 1
                            Else

                                GoTo 100
                            End If

                        Case Else
                            If zkBgArr(zkBgArr.Length - 1).dsd + bgspace + 0.3 < tcTbl(i).tcdsd Then
                                If Not (zkBgArr.Length = 1 And zkBgArr(0).bggc = 0) Then ReDim Preserve zkBgArr(zkBgArr.Length)
                                zkBgArr(zkBgArr.Length - 1).dsd = zkBgArr(zkBgArr.Length - 2).dsd + bgspace
                                zkBgArr(zkBgArr.Length - 1).bggc = zkBgArr(zkBgArr.Length - 1).dsd + 0.8
                                zkBgArr(zkBgArr.Length - 1).bgjs = getRndInt(dtbl(getIndexInDgvTbl(dtbl, tcTbl(i).tcbh)).loLmt, dtbl(getIndexInDgvTbl(dtbl, tcTbl(i).tcbh)).UpLmt)
                                m = m + 1
                                dtbl(getIndexInDgvTbl(dtbl, tcTbl(i).tcbh)).testCount = dtbl(getIndexInDgvTbl(dtbl, tcTbl(i).tcbh)).testCount - 1
                            Else
                                Exit Do
                            End If
                    End Select


                Loop


            End If
100:    Next
    End Sub
    Sub getTcBgarr(ByRef dtbl() As typeDgvTbl, ByRef tcBgarr() As TypeBG)

    End Sub

    Sub getZKDtArr(ByRef zkDtArr() As TypeDt, ByVal tcTbl() As typetc, ByVal dtbl() As typeDgvTbl)
        For i As Integer = 0 To tcTbl.Length - 1
            Dim tcNarr() As Double
            If dtbl(getIndexInDgvTbl(dtbl, tcTbl(i).tcbh)).UpLmt > 0 Then
                If tcTbl(i).tclen > 0.1 Then
                    getTcNarr(getRndN(dtbl(getIndexInDgvTbl(dtbl, tcTbl(i).tcbh)).loLmt, dtbl(getIndexInDgvTbl(dtbl, tcTbl(i).tcbh)).UpLmt), tcNarr, tcTbl(i).tclen, dtbl, getIndexInDgvTbl(dtbl, tcTbl(i).tcbh))
                    fillZkDtArr(zkDtArr, tcNarr, dtbl(getIndexInDgvTbl(dtbl, tcTbl(i).tcbh)).cslx, dtbl(getIndexInDgvTbl(dtbl, tcTbl(i).tcbh)).cv)
                End If

            Else
                If zkDtArr.Length = 0 And zkDtArr(0).dsd = 0 Then
                    ReDim Preserve zkDtArr(System.Math.Round(tcTbl(i).tclen / UnitLen, 0) - 1)
                    For n As Integer = 0 To zkDtArr.Length - 1
                        If n = 0 Then
                            zkDtArr(n).dsd = UnitLen
                            zkDtArr(n).dtlx = eDtlx.N10
                        Else
                            zkDtArr(n).dsd = zkDtArr(n - 1).dsd + UnitLen
                            zkDtArr(n).dtlx = eDtlx.N10
                        End If
                    Next
                    GoTo 10
                End If
                If zkDtArr.Length <> 0 Then
                    Dim lastlen As Integer = zkDtArr.Length
                    ReDim Preserve zkDtArr(zkDtArr.Length + System.Math.Round(tcTbl(i).tclen / UnitLen, 0) - 1)
                    For n As Integer = lastlen To zkDtArr.Length - 1
                        zkDtArr(n).dsd = zkDtArr(n - 1).dsd + UnitLen
                        zkDtArr(n).dtlx = eDtlx.N10
                    Next
                End If

            End If

10:     Next

    End Sub

    Sub fillZkDtArr(ByRef zkDtArr() As TypeDt, ByVal tcNarr() As Double, ByVal dtlx As eDtlx, ByVal cv As Double)
        If zkDtArr.Length = 1 And zkDtArr(0).dtgc = 0 Then
            ReDim zkDtArr(tcNarr.Length - 1)
            For i As Integer = 0 To zkDtArr.Length - 1
                With zkDtArr(i)
                    .dtjs = tcNarr(i)
                    .dsd = (i + 1) * UnitLen
                    .dtgc = zkDtArr.Length * UnitLen + 0.5
                    .dtlx = dtlx
                End With

            Next
            Exit Sub
        Else
            Dim lastLen As Integer = zkDtArr.Length
            Dim lastGc As Double = zkDtArr(zkDtArr.Length - 1).dtgc
            Dim lastDSD As Double = zkDtArr(zkDtArr.Length - 1).dsd
            Dim lastDtlx As eDtlx = zkDtArr(zkDtArr.Length - 1).dtlx
            Dim lastDtjs As Double = zkDtArr(zkDtArr.Length - 1).dtjs
            ReDim Preserve zkDtArr(lastLen + tcNarr.Length - 1)

            If lastDtlx = dtlx And lastLen > 2 And tcNarr.Length > 2 And System.Math.Abs(lastDtjs - tcNarr(0)) > 50 * cv Then
                If lastDtjs > tcNarr(0) Then

                    zkDtArr(zkDtArr.Length - tcNarr.Length - 2).dtjs = lastDtjs
                    zkDtArr(zkDtArr.Length - tcNarr.Length - 1).dtjs = lastDtjs - (lastDtjs - tcNarr(0)) / 3
                    With zkDtArr(zkDtArr.Length - tcNarr.Length)
                        .dtjs = lastDtjs - 2 * (lastDtjs - tcNarr(0)) / 3
                        .dsd = lastDSD + UnitLen * 1
                        .dtgc = zkDtArr.Length * UnitLen + 0.5
                        .dtlx = dtlx
                    End With
                    With zkDtArr(zkDtArr.Length - tcNarr.Length + 1)
                        .dtjs = tcNarr(0)
                        .dsd = lastDSD + UnitLen * 2
                        .dtgc = zkDtArr.Length * UnitLen + 0.5
                        .dtlx = dtlx
                    End With


                Else
                    zkDtArr(zkDtArr.Length - tcNarr.Length - 2).dtjs = lastDtjs
                    zkDtArr(zkDtArr.Length - tcNarr.Length - 1).dtjs = lastDtjs + (tcNarr(0) - lastDtjs) / 3
                    With zkDtArr(zkDtArr.Length - tcNarr.Length)
                        .dtjs = lastDtjs + 2 * (tcNarr(0) - lastDtjs) / 3
                        .dsd = lastDSD + UnitLen * 1
                        .dtgc = zkDtArr.Length * UnitLen + 0.5
                        .dtlx = dtlx
                    End With
                    With zkDtArr(zkDtArr.Length - tcNarr.Length + 1)
                        .dtjs = tcNarr(0)
                        .dsd = lastDSD + UnitLen * 2
                        .dtgc = zkDtArr.Length * UnitLen + 0.5
                        .dtlx = dtlx
                    End With

                End If
                For i As Integer = 2 To tcNarr.Length - 1
                    With zkDtArr(lastLen + i)
                        .dtjs = tcNarr(i)
                        .dsd = lastDSD + (i + 1) * UnitLen
                        .dtgc = zkDtArr.Length * UnitLen + 0.5
                        .dtlx = dtlx
                    End With
                Next
            Else
                For i As Integer = 0 To tcNarr.Length - 1
                    With zkDtArr(lastLen + i)
                        .dtjs = tcNarr(i)
                        .dsd = lastDSD + (i + 1) * UnitLen
                        .dtgc = zkDtArr.Length * UnitLen + 0.5
                        .dtlx = dtlx
                    End With
                Next
            End If



        End If

    End Sub
    Function getIndexInDgvTbl(ByVal dtbl() As typeDgvTbl, ByVal tcbh As String) As Integer
        For i As Integer = 0 To dtbl.Length - 1
            If dtbl(i).tcbh = tcbh Then
                Return i

            End If
        Next
    End Function
    Sub getTcNarr(ByVal sN As Double, ByRef TcNarr() As Double, ByVal tcLen As Double, ByVal dTbl() As typeDgvTbl, ByVal index As Integer)
        Try
            Dim dr As eDrection = getThisDr(eDrection.none)
            Dim lb As Double = dTbl(index).loLmt
            Dim ub As Double = dTbl(index).UpLmt
            Dim cv As Double = dTbl(index).cv
            Dim minL As Double = dTbl(index).MinLenght
            ReDim TcNarr(0)
            TcNarr(0) = sN

            Do While True
                Dim LastingCount As Integer = getRndIntN(3, 10)
                Dim lastd As Double = TcNarr(TcNarr.Length - 1)
                Dim lastTcNarrUbound As Integer = TcNarr.Length - 1
                If (TcNarr.Length + LastingCount) * UnitLen > tcLen Then

                    ReDim Preserve TcNarr(System.Math.Round(tcLen / UnitLen, 0) - 1)
                    LastingCount = TcNarr.Length - lastTcNarrUbound - 1
                    Dim Narr() As Double
                    getRndNarrByDr(tcLen, minL, lastd, dr, ub, lb, cv, LastingCount, Narr)
                    dr = getThisDr(dr)
                    For i As Integer = 1 To LastingCount
                        TcNarr(lastTcNarrUbound + i) = Narr(i - 1)
                    Next
                    Exit Do
                Else
                    ReDim Preserve TcNarr(TcNarr.Length + LastingCount - 1)
                    Dim Narr() As Double
                    getRndNarrByDr(tcLen, minL, lastd, dr, ub, lb, cv, LastingCount, Narr)
                    dr = getThisDr(dr)
                    For i As Integer = 1 To LastingCount
                        TcNarr(lastTcNarrUbound + i) = Narr(i - 1)
                    Next
                End If
            Loop
        Catch ex As Exception
            MsgBox("gettcnarr中出错-" & Err.Description & Err.Erl)
        End Try


    End Sub

    Sub scrub(ByRef zkdtarr() As TypeDt, ByVal scrubCount As Integer, ByVal scrubInterval As Integer)
        scrubCount = Int(scrubCount)
        scrubInterval = Int(scrubInterval)
        If scrubCount < 0 Or scrubInterval <= 0 Then Exit Sub

        If zkdtarr.Length > 30 Then

            Dim l As Integer
            Do While True
                scrubInterval = scrubInterval + getRndInt(0, 2)
                scrubCount = scrubCount + getRndInt(0, 2)
                l = l + scrubCount + scrubInterval
                If l > zkdtarr.Length Then Exit Sub
                For i As Integer = l - scrubCount - 1 To l - 1
                    zkdtarr(i).dtjs = 0
                Next
            Loop
        End If
    End Sub
    Sub getRndNarrByDr(ByVal tclen As Double, ByVal minLen As Double, ByVal lastD As Double, ByVal dr As eDrection, ByVal ub As Double, ByVal lb As Double, ByVal cv As Double, ByVal count As Integer, ByRef Narr() As Double)
        ReDim Narr(count - 1)
        If tclen < minLen Then Exit Sub
        For i As Integer = 0 To Narr.Length - 1
            If i = 0 Then
                Select Case dr
                    Case eDrection.decrease
                        Narr(i) = getRndN(LowLmt(lb, ub, cv), lastD, lastD, cv)
                    Case eDrection.increase
                        Narr(i) = getRndN(lastD, UpLmt(lb, ub, cv), lastD, cv)
                    Case eDrection.none
                        Narr(i) = getRndN(LowLmt(lastD, lastD, cv), UpLmt(lastD, lastD, cv))
                End Select
            Else
                Select Case dr
                    Case eDrection.decrease
                        Narr(i) = getRndN(LowLmt(lb, ub, cv), Narr(i - 1), Narr(i - 1), cv)
                    Case eDrection.increase
                        Narr(i) = getRndN(Narr(i - 1), UpLmt(lb, ub, cv), Narr(i - 1), cv)
                    Case eDrection.none
                        Narr(i) = getRndN(LowLmt(Narr(i - 1), Narr(i - 1), cv), UpLmt(Narr(i - 1), Narr(i - 1), cv))
                End Select
            End If
        Next

    End Sub


    Public Sub New()
        UnitLen = 0.1
    End Sub
End Class
