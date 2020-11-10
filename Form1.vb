Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic

Imports ADODB
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports System.Drawing
Imports System.IO
Imports System.Math

Imports MSWord = Microsoft.Office.Interop.Word

Friend Class Form1
    Inherits System.Windows.Forms.Form
    Public PubCurrentUserName当前用户 As String

    Enum FHCD
        全风化 = 1
        强风化 = 2
        中风化 = 3
        微风化 = 4
    End Enum



    Public cnt As New ADODB.Connection
    '全局变量
    Dim gcsy As Integer
    Public PubArr土层分类() As QyclsTcAlteration.TypeTuceng
    Public PubList本工程钻孔表集合 As New List(Of 钻孔编号高程数据结构)
    Public PubList本工程地层数据集合 As New List(Of 地层数据结构)
    '动探，标贯，取样相关钻孔数组似乎没有作用
    Dim dtRelatedZKs() As qyywcsOptor.relatedZK
    Dim bgRelatedZKs() As qyywcsOptor.relatedZK
    Dim qyRelatedZKs() As qyywcsOptor.relatedZK

    Dim 超重型动力触探修正表集合 As List(Of String)
    Dim 重型动力触探修正表集合 As List(Of String)

    Public Sub New()

        ' 此调用是设计器所必需的。
        InitializeComponent()

        ' 在 InitializeComponent() 调用之后添加任何初始化。

    End Sub

    Private Sub UserForm1_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        Try
            'usingRecorder.endTime = Now.ToString
            'usingRecorder.addEndTime()
        Catch ex As Exception
            MsgBox(Err.Description)
        End Try



    End Sub
    '定义RC1，RC2,用来构造包含钻孔一和钻孔二包含土层名称及其延伸类型，贯通土层等信息的表

    Private Sub UserForm_Initialize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Try


            超重型动力触探修正表集合 = New List(Of String)
            超重型动力触探修正表集合.Clear()
            Dim s() As String = File.ReadAllLines(Application.StartupPath & "\超重型动力触探修正表.txt")
            For i As Integer = 0 To UBound(s)
                超重型动力触探修正表集合.Add(s(i))
            Next

            重型动力触探修正表集合 = New List(Of String)
            重型动力触探修正表集合.Clear()
            s = File.ReadAllLines(Application.StartupPath & "\重型动力触探修正表.txt")
            For i As Integer = 0 To UBound(s)
                重型动力触探修正表集合.Add(s(i))
            Next
            dgvchanzhuang.RowCount = 1
        Catch ex As Exception
            MsgBox("formload中-" & Err.Description)
        End Try



    End Sub
    Private Sub DataGridView1_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles dgvDt.CellEnter, dgvquyang.CellEnter, dgvbg.CellEnter, dgvfenghua.CellEnter
        If TypeOf (CType(sender, DataGridView).Columns(e.ColumnIndex)) Is DataGridViewComboBoxColumn Then
            SendKeys.Send("{F4}")
            Exit Sub
        End If
    End Sub
    Private Sub OpenDataBase_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OpenDataBase.Click
        Try

            '打开理正数据库并连接它
            Dim Filedatabase As String
            If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                Filedatabase = OpenFileDialog1.FileName
            End If
            If Filedatabase = "" Then
                Exit Sub
            End If
            '连接理正数据库
            If cnt.ConnectionString <> "" Then
                MsgBox("数据库已经打开。请点击关闭数据库再连接其他数据库")
                Exit Sub
            End If

            cnt.Open("provider=microsoft.jet.oledb.4.0;data source=" & Filedatabase)
            lbl文件.Text = Filedatabase
            func读出钻孔表()
            func读出地层表并计算土层厚度和层顶底标高(lblStatus)
            updateDgv()

        Catch ex As Exception
            MsgBox("opendatabase过程中出现错误-" & Err.Description)
        End Try

    End Sub
    Sub updateDgv()
        Try
            dgvfenghua.Rows.Clear()
            dgvTcAlter.Rows.Clear()
            dgvDt.Rows.Clear()
            dgvbg.Rows.Clear()
            dgvquyang.Rows.Clear()

            If Func从数据库中获取土层分类数组(PubArr土层分类) = False Then
                MsgBox("从工程数据库中获取土层名称失败。")
                Exit Sub
            End If
            dgvfenghua.RowCount = PubArr土层分类.Length
            dgvDt.RowCount = PubArr土层分类.Length
            dgvbg.RowCount = PubArr土层分类.Length
            dgvquyang.RowCount = PubArr土层分类.Length
            dgvTcAlter.RowCount = PubArr土层分类.Length
            dgvbosu.RowCount = PubArr土层分类.Length

            Dim f As FHCD
            Dim o As Object = System.Enum.GetValues(f.GetType)
            Dim a(UBound(o)) As String
            For i As Integer = 0 To UBound(o)
                a(i) = o(i).ToString
            Next
            Dim y As qyywcsOptor.eDtlx
            o = System.Enum.GetValues(y.GetType)
            Dim b(UBound(o)) As String
            For i As Integer = 0 To UBound(o)
                b(i) = o(i).ToString
            Next

            ColFHDJ.DataSource = a
            ColDtlx.DataSource = b



            For i As Integer = 0 To PubArr土层分类.Length - 1
                dgvfenghua.Item(0, i).Value = PubArr土层分类(i).tcbh
                dgvfenghua.Item(1, i).Value = PubArr土层分类(i).tcmc
                dgvTcAlter.Item(0, i).Value = PubArr土层分类(i).tcbh
                dgvTcAlter.Item(1, i).Value = PubArr土层分类(i).tcmc
                dgvDt.Item(0, i).Value = PubArr土层分类(i).tcbh
                dgvDt.Item(1, i).Value = PubArr土层分类(i).tcmc
                dgvbg.Item(0, i).Value = PubArr土层分类(i).tcbh
                dgvbg.Item(1, i).Value = PubArr土层分类(i).tcmc
                dgvquyang.Item(0, i).Value = PubArr土层分类(i).tcbh
                dgvquyang.Item(1, i).Value = PubArr土层分类(i).tcmc
                dgvbosu.Item(0, i).Value = PubArr土层分类(i).tcbh
                dgvbosu.Item(1, i).Value = PubArr土层分类(i).tcmc
            Next

        Catch ex As Exception
            MsgBox(Err.Description)
        End Try

    End Sub


    Function Func从数据库中获取土层分类数组(ByRef PubArr土层分类() As QyclsTcAlteration.TypeTuceng) As Boolean
        Try

            ReDim PubArr土层分类(0)
            Dim rs As New ADODB.Recordset
            rs.Open("select * from z_g_TuCeng order by tczcbh,tcycbh", cnt, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly)
            If rs.RecordCount > 0 Then
                rs.MoveFirst()
                Do Until rs.EOF
                    Dim bh As String = rs.Fields("tczcbh").Value & "-" & rs.Fields("tcycbh").Value

                    Dim mc As String
                    Try
                        mc = rs.Fields("tcmc").Value
                    Catch ex As Exception
                        MsgBox("错误，" & rs.Fields("zkbh").Value & "土层名称字段为空")
                        Return False
                    End Try

                    If PubArr土层分类(0).tcbh = "" Then
                        PubArr土层分类(0).tcbh = bh
                        PubArr土层分类(0).tcmc = mc
                    ElseIf ifIntcbhmcarr(PubArr土层分类, bh) = False Then
                        ReDim Preserve PubArr土层分类(PubArr土层分类.Length)
                        PubArr土层分类(PubArr土层分类.Length - 1).tcbh = bh
                        PubArr土层分类(PubArr土层分类.Length - 1).tcmc = mc

                    End If
                    rs.MoveNext()
                Loop
            End If
            rs.Close()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)
            If PubArr土层分类(0).tcmc <> "" Then
                Return True
            End If
        Catch ex As Exception
            MsgBox("gettcbhtcmcfromrs中发生错误," & Err.Description)
            Return False
        End Try

    End Function
    Function ifIntcbhmcarr(ByVal tcbhmcarr() As QyclsTcAlteration.TypeTuceng, ByVal tcbh As String) As Boolean
        For i As Integer = 0 To tcbhmcarr.Length - 1
            If tcbhmcarr(i).tcbh = tcbh Then
                Return True
                Exit Function
            End If
        Next
    End Function

    Private Sub CloseConnection_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CloseConnection.Click
        Try
            cnt.Close()
            cnt.ConnectionString = ""
            lbl文件.Text = ""
            'List1.Items.Clear()



            dgvfenghua.Rows.Clear()
            dgvchanzhuang.Rows.Clear()
            dgvTcAlter.Rows.Clear()

            dgvDt.Rows.Clear()
            dgvbg.Rows.Clear()
            dgvquyang.Rows.Clear()
            dgvbosu.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub UserForm_Terminate()
        cnt.Close()
    End Sub


















    Private Sub dgvfenghua_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvfenghua.DataError
        MsgBox(e.Exception.Message)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MsgBox(dgvfenghua.Item(2, 0).Value)
    End Sub

    Private Sub dgvywcs_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvDt.CellValueChanged
        Try
            If e.ColumnIndex = 0 Then
                For i As Integer = 0 To PubArr土层分类.Length - 1
                    If dgvDt.Item(0, e.RowIndex).Value = PubArr土层分类(i).tcbh Then
                        dgvDt.Item(1, e.RowIndex).Value = PubArr土层分类(i).tcmc
                        Exit Sub
                    End If
                Next
            End If
        Catch ex As Exception
            'MsgBox(Err.Description)
        End Try


    End Sub

    Private Sub bntruku_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntruku.Click
        Try

            Dim cmd As New ADODB.Command
            cmd.ActiveConnection = cnt
            For i As Integer = 0 To dgvfenghua.RowCount - 1
                Dim t As Object = VB.Split(dgvfenghua.Item(0, i).Value, "-", -1, CompareMethod.Text)
                If dgvfenghua.Item(2, i).Value <> "" Then
                    cmd.CommandText = "update z_g_tuceng set tcfhcd='" & dgvfenghua.Item(2, i).Value & "' where tczcbh=" & Val(t(0)) & " and tcycbh=" & Val(t(1))
                    cmd.Execute()
                    MsgBox("风化数据已入库")
                End If

            Next
            '读取倾向倾角
            If IsNumeric(dgvchanzhuang.Item(0, 0).Value) = False Or IsNumeric(dgvchanzhuang.Item(1, 0).Value) = False Then
                MsgBox("倾角或倾向不为数字。")
                Exit Sub
            Else
                Dim r As New Recordset
                'r.Open("select * from z_g_tuceng where tcmc LIKE '%岩%'", cnt, CursorTypeEnum.adOpenStatic)
                'MsgBox(r.RecordCount)
                Dim cmdStr As String
                cmdStr = "update z_g_tuceng set tcysqx=" & dgvchanzhuang.Item(0, 0).Value & ",tcysqj=" & dgvchanzhuang.Item(1, 0).Value & " where  TCLM like '%岩%'"
                cmd.CommandText = cmdStr
                cmd.Execute()

            End If

            MsgBox("产状数据已入库。")

        Catch ex As Exception
            MsgBox("入库块中-" & Err.Description)

        End Try


    End Sub
    Function FuncSort地层集合按钻孔编号和层顶深度排序函数(x As 地层数据结构, y As 地层数据结构) As Integer
        If x.zkbh > y.zkbh Then
            Return 1
        ElseIf x.zkbh = y.zkbh Then
            If x.cengdishengdu > y.cengdishengdu Then
                Return 1
            ElseIf x.cengdishengdu = y.cengdishengdu Then
                Return 0
            Else
                Return -1
            End If
        Else
            Return -1
        End If
    End Function
    Function func读出钻孔表() As Boolean
        Try
            PubList本工程钻孔表集合.Clear()
            Dim rsTuceng As New Recordset
            rsTuceng.Open("Select * from z_zuankong", cnt, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly)
            If rsTuceng.RecordCount = 0 Then
                Return True
            End If
            rsTuceng.MoveFirst()
            Try
                If IsDBNull(rsTuceng.Fields("gcsy").Value) = False Then gcsy = rsTuceng.Fields("gcsy").Value
            Catch ex As Exception
                MsgBox("数据库中z_zuankong表中gcsy字段未填写")
                Return False
            End Try

            '读出数据
            Do Until rsTuceng.EOF
                Dim t As 钻孔编号高程数据结构
                With t
                    Try
                        .zkbh = rsTuceng.Fields("zkbh").Value
                        .zkbg = rsTuceng.Fields("zkbg").Value
                        .zkx = rsTuceng.Fields("zkx").Value
                        .zky = rsTuceng.Fields("zky").Value

                        'added by Nans @ 20201108

                        If TypeName(rsTuceng.Fields("ZKKSRQ").Value) <> "DBNull" Then .ZKKSRQ = rsTuceng.Fields("ZKKSRQ").Value Else .ZKKSRQ = "20201108" '钻孔开始日期

                        If TypeName(rsTuceng.Fields("ZKZZRQ").Value) <> "DBNull" Then .ZKZZRQ = rsTuceng.Fields("ZKZZRQ").Value Else .ZKZZRQ = "20201115" '钻孔终止日期

                        If TypeName(rsTuceng.Fields("ZKZJLX").Value) <> "DBNull" Then .ZKZJLX = rsTuceng.Fields("ZKZJLX").Value Else .ZKZJLX = "NULL" '钻机类型

                    Catch ex As Exception
                        MsgBox("读出钻孔" & .zkbh & "出错")
                        Return False
                    End Try
                End With
                PubList本工程钻孔表集合.Add(t)
                rsTuceng.MoveNext()
            Loop
        Catch ex As Exception
            MsgBox("读出钻孔表出错" & Err.Description)
            Return False
        End Try
    End Function
    Public Function func读出地层表并计算土层厚度和层顶底标高(ByRef lblstatus As System.Windows.Forms.Label) As Boolean
        'On Error Resume Next

        PubList本工程地层数据集合.Clear()
        Try
            Dim rsTuceng As New Recordset
            rsTuceng.Open("Select * from z_g_TuCeng order by zkbh, tccdsd", cnt, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly)


            If rsTuceng.RecordCount = 0 Then
                Return True
            End If
            rsTuceng.MoveFirst()
            '读出数据
            Do Until rsTuceng.EOF
                Dim t As New 地层数据结构(0)
                lblstatus.Text = "正在读取" & rsTuceng.Fields("zkbh").Value
                lblstatus.Refresh()
                With t
                    .zkbh = rsTuceng.Fields("zkbh").Value
                    Dim d As Double = PubList本工程钻孔表集合.Find(Function(s As 钻孔编号高程数据结构) s.zkbh = .zkbh).zkbg
                    .zkbg = System.Math.Round(d, 3)
                    .TCZCBH = rsTuceng.Fields("tczcbh").Value
                    .TCYCBH = rsTuceng.Fields("tcycbh").Value
                    .tccycbh = rsTuceng.Fields("tccycbh").Value
                    .cengdishengdu = System.Math.Round(rsTuceng.Fields("tccdsd").Value, 2)
                    .cengdibiaogao = System.Math.Round(.zkbg - .cengdishengdu, 2)

                    'ADDED BY NANS @20201108
                    If TypeName(rsTuceng.Fields("TCYS").Value) <> "DBNull" Then .TCYS = rsTuceng.Fields("TCYS").Value Else .TCYS = "NULL" '土层颜色
                    If TypeName(rsTuceng.Fields("TCKSX").Value) <> "DBNull" Then .TCKSX = rsTuceng.Fields("TCKSX").Value Else .TCKSX = "NULL" '土层可塑性
                    If TypeName(rsTuceng.Fields("TCSID").Value) <> "DBNull" Then .TCSID = rsTuceng.Fields("TCSID").Value Else .TCSID = "NULL" '土层可塑性
                    If TypeName(rsTuceng.Fields("TCMS").Value) <> "DBNull" Then .TCMS = rsTuceng.Fields("TCMS").Value Else .TCMS = "NULL" '土层可塑性
                    If TypeName(rsTuceng.Fields("TCMSD").Value) <> "DBNull" Then .TCMSD = rsTuceng.Fields("TCMSD").Value Else .TCMSD = "NULL" '土层可塑性

                    If IsDBNull(rsTuceng.Fields("tcmc").Value) = False Then
                        .tcmc = rsTuceng.Fields("tcmc").Value
                    Else
                        MsgBox(.zkbh & "中土层名称为空")
                        Return False
                    End If
                    If IsDBNull(rsTuceng.Fields("tclm").Value) = False Then .tclm = rsTuceng.Fields("tclm").Value

                End With
                PubList本工程地层数据集合.Add(t)
                rsTuceng.MoveNext()
            Loop
            '计算各记录的土层厚度，层顶标高，层底标高等
            PubList本工程地层数据集合.Sort(AddressOf FuncSort地层集合按钻孔编号和层顶深度排序函数)

            Dim a() As 地层数据结构 = PubList本工程地层数据集合.ToArray
            For i As Integer = 0 To a.Length - 1
                With a(i)
                    If i = 0 Then
                        .tchd = .cengdishengdu
                        .cengdingshendu = .cengdishengdu - .tchd
                        .cengdingbiaogao = .cengdibiaogao + .tchd
                    ElseIf a(i).zkbh = a(i - 1).zkbh Then
                        .tchd = .cengdishengdu - a(i - 1).cengdishengdu
                        .cengdingshendu = .cengdishengdu - .tchd
                        .cengdingbiaogao = .cengdibiaogao + .tchd
                    Else
                        .tchd = .cengdishengdu
                        .cengdingshendu = .cengdishengdu - .tchd
                        .cengdingbiaogao = .cengdibiaogao + .tchd
                    End If
                End With
            Next
            PubList本工程地层数据集合 = a.ToList

            lblstatus.Text = ""
            lblstatus.Refresh()
            Return True
        Catch ex As Exception
            MsgBox("读出地层表过程出错" & Err.Description)
            Return False
        End Try

    End Function


    Function 生成波速取芯率和RQD() As Boolean
        Try
            '读控件表

            publst波速表土层设置.Clear()
            publst取芯率RQD设置表.Clear()

            Dim d As DataGridView = dgvbosu
            Dim tcbh As String
            Dim tcmc As String
            Dim hengbololmt As Double
            Dim hengbouplmt As Double
            Dim hengbocv As Double
            Dim zongbololmt As Double
            Dim zongbouplmt As Double
            Dim zongbocv As Double
            Dim qxllo As Double
            Dim qxlup As Double

            Dim rqdlo As Double
            Dim rqdup As Double
            '读取波速设置表，生成波速表规则集合
            For i As Integer = 0 To d.RowCount - 1

                Try
                    tcbh = d.Item(0, i).Value
                    tcmc = d.Item(1, i).Value

                    hengbololmt = Split(d.Item(2, i).Value, "-")(0)
                    hengbouplmt = Split(d.Item(2, i).Value, "-")(1)
                    'hengbocv = d.Item(3, i).Value
                    zongbololmt = Split(d.Item(4, i).Value, "-")(0)
                    zongbouplmt = Split(d.Item(4, i).Value, "-")(1)
                    'zongbocv = d.Item(5, i).Value
                    'qxllo = Split(d.Item(6, i).Value, "-")(0)
                    'qxlup = Split(d.Item(6, i).Value, "-")(1)
                    'rqdlo = Split(d.Item(7, i).Value, "-")(0)
                    'rqdup = Split(d.Item(7, i).Value, "-")(1)

                    If hengbololmt <= 0 Or hengbouplmt <= 0 Or zongbololmt <= 0 Or zongbouplmt <= 0 Then
                        MsgBox("控件表中参数小于0不满足要求")
                        Return False
                    End If

                    Dim success As Boolean
                    Dim t As New struDgv波速表设置(tcbh, tcmc, hengbololmt, hengbouplmt, hengbocv, zongbololmt, zongbouplmt, zongbocv, success)
                    publst波速表土层设置.Add(t)
                Catch ex As Exception
                    'MsgBox("波速表设置中参数设置不符合要求，请检查参数是否为数字" & Err.Description)
                    '                    Return False
                End Try

            Next
            '读取取芯率和RQD设置表，生成取芯率和RQD规则集合
            For i As Integer = 0 To d.RowCount - 1
                Try

                    If d.Item(6, i).Value <> "" Or d.Item(7, i).Value <> "" Then
                        tcbh = d.Item(0, i).Value
                        tcmc = d.Item(1, i).Value
                        If d.Item(6, i).Value <> "" Then
                            If UBound(Split(d.Item(6, i).Value, "-")) <> 1 Then
                                MsgBox("取芯率范围值设置不满足要求，正确格式为两个大于0小于100的数字中间用符合'-'隔开,且左边数字小于右边数字，如20-30")
                            Else
                                If IsNumeric(Split(d.Item(6, i).Value, "-")(0)) = False Or IsNumeric(Split(d.Item(6, i).Value, "-")(1)) = False Then
                                Else
                                    If Split(d.Item(6, i).Value, "-")(0) <= 0 Or Split(d.Item(6, i).Value, "-")(1) <= 0 Then
                                    Else
                                        If Split(d.Item(6, i).Value, "-")(0) >= Split(d.Item(6, i).Value, "-")(1) Then
                                        Else
                                            If Split(d.Item(6, i).Value, "-")(0) > 100 Then
                                            Else
                                                '满足要求的数据
                                                qxllo = Split(d.Item(6, i).Value, "-")(0)
                                                qxlup = Split(d.Item(6, i).Value, "-")(1)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        If d.Item(7, i).Value <> "" Then
                            If UBound(Split(d.Item(7, i).Value, "-")) <> 1 Then
                                MsgBox("RQD范围值设置不满足要求，正确格式为两个大于0小于100的数字中间用符合'-'隔开,且左边数字小于右边数字，如20-30")
                            Else
                                If IsNumeric(Split(d.Item(7, i).Value, "-")(0)) = False Or IsNumeric(Split(d.Item(7, i).Value, "-")(1)) = False Then
                                Else
                                    If Split(d.Item(7, i).Value, "-")(0) <= 0 Or Split(d.Item(7, i).Value, "-")(1) <= 0 Then
                                    Else
                                        If Split(d.Item(7, i).Value, "-")(0) >= Split(d.Item(7, i).Value, "-")(1) Then
                                        Else
                                            If Split(d.Item(7, i).Value, "-")(0) > 100 Then
                                            Else
                                                '满足要求的数据
                                                rqdlo = Split(d.Item(7, i).Value, "-")(0)
                                                rqdup = Split(d.Item(7, i).Value, "-")(1)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If

                    Dim t As New struDgv取芯和RQD(tcbh, tcmc, qxllo, qxlup, rqdlo, rqdup)
                    publst取芯率RQD设置表.Add(t)
                Catch ex As Exception

                End Try
            Next



            '生成相关钻孔集合
            Dim bosuzks As New List(Of String）
            bosuzks.Clear()
            bosuzks = Split(txtbxbosuspecifiedZks.Text, ",").ToList

            Dim cmd As New ADODB.Command
            cmd.ActiveConnection = cnt
            cmd.CommandText = "delete * from z_y_bosu"
            Try
                cmd.Execute()
            Catch ex As Exception
                MsgBox("删除波速记录出错")
                Return False
            End Try
            cmd.CommandText = "delete * from z_g_yxrqd"
            Try
                cmd.Execute()
            Catch ex As Exception
                MsgBox("删除取芯RQD记录出错")
                Return False
            End Try
            '生成每个钻孔数据
            For i As Integer = 0 To bosuzks.Count - 1
                Dim lstzktc As New List(Of 地层数据结构）
                lstzktc.Clear()
                lstzktc = PubList本工程地层数据集合.FindAll(Function(ss As 地层数据结构) ss.zkbh = bosuzks(i))
                Dim lstbosu As New List(Of struBosu)
                lstbosu.Clear()
                Dim lstqxlrqd As New List(Of stru取芯RQD）
                lstqxlrqd.Clear()

                lstbosu = funcGetZkBosuLst(lstzktc)
                lstqxlrqd = funcGetZkQxlRqdLst(lstzktc)

                lblStatus.Text = "正在对" & bosuzks(i) & "添加波速数据"
                lblStatus.Refresh()
                For ii As Integer = 0 To lstbosu.Count - 1
                    cmd.CommandText = "insert into z_y_bosu (zkbh,gcsy,bssd,bshbs,bszbs,bsshb,bsszb) values('" & bosuzks(i) & "'," & gcsy & "," & lstbosu(ii).dsd & "," & lstbosu(ii).bosuh & "," & lstbosu(ii).bosuz & ",1,1)"
                    Try
                        cmd.Execute()
                    Catch ex As Exception
                        MsgBox(bosuzks(i) & "添加波速数据出错" & Err.Description)
                        Return False
                    End Try
                Next
                lblStatus.Text = "正在对" & bosuzks(i) & "添加取芯率和RQD数据"
                lblStatus.Refresh()
                For ii As Integer = 0 To lstqxlrqd.Count - 1
                    If lstqxlrqd(ii).qxl = 0 Then
                        cmd.CommandText = "insert into z_g_yxrqd (zkbh,gcsy,yrsd,yrrqd,yrsrqd) values('" & bosuzks(i) & "'," & gcsy & "," & lstqxlrqd(ii).dsd & "," & lstqxlrqd(ii).rqd & ",1)"
                    ElseIf lstqxlrqd(ii).rqd = 0 Then
                        cmd.CommandText = "insert into z_g_yxrqd (zkbh,gcsy,yrsd,yrcql,yrscql) values('" & bosuzks(i) & "'," & gcsy & "," & lstqxlrqd(ii).dsd & "," & lstqxlrqd(ii).qxl & ",1)"
                    Else
                        cmd.CommandText = "insert into z_g_yxrqd (zkbh,gcsy,yrsd,yrcql,yrrqd,yrscql,yrsrqd) values('" & bosuzks(i) & "'," & gcsy & "," & lstqxlrqd(ii).dsd & "," & lstqxlrqd(ii).qxl & "," & lstqxlrqd(ii).rqd & ",1,1)"
                    End If


                    Try
                        cmd.Execute()
                    Catch ex As Exception
                        MsgBox(bosuzks(i) & "添加取芯率RQD数据出错" & Err.Description)
                        Return False
                    End Try
                Next
            Next
            '添加的数据库

        Catch ex As Exception
            MsgBox("波速或取芯RQD记录出错")
            Return False
        End Try
    End Function
    Public publst波速表土层设置 As New List(Of struDgv波速表设置)
    Public publst取芯率RQD设置表 As New List(Of struDgv取芯和RQD）
    Public Function funcGetZkQxlRqdLst(zktclst As List(Of 地层数据结构)) As List(Of stru取芯RQD)
        Dim lstqxlrqd As New List(Of stru取芯RQD)
        lstqxlrqd.Clear()
        Try

            Dim dsd As Double

            Dim interval As Double
            Try
                interval = tbx钻探回次长度.Text
                If interval <= 0 Then
                    MsgBox("回次长度不能小于等于0")
                    Return lstqxlrqd
                End If
            Catch ex As Exception
                MsgBox("回次长度必须为数字")
                Return lstqxlrqd
            End Try
            Try
                dsd = tbx钻探回次长度.Text
                If dsd <= 0 Then
                    MsgBox("回次长度不能小于等于0")
                    Return lstqxlrqd
                End If
            Catch ex As Exception
                MsgBox("回次长度必须为数字")
                Return lstqxlrqd
            End Try

            Do While True

                Dim tc As 地层数据结构
                Dim qxl As Double
                Dim rqd As Double
                If zktclst.Exists(Function(ss As 地层数据结构) ss.cengdingshendu < dsd And ss.cengdishengdu >= dsd) Then
                    tc = zktclst.Find(Function(ss As 地层数据结构) ss.cengdingshendu < dsd And ss.cengdishengdu >= dsd)
                    Dim tcbh As String
                    tcbh = tc.TCZCBH & "-" & tc.TCYCBH
                    Dim lo As Double
                    Dim up As Double
                    If publst取芯率RQD设置表.Exists(Function(ss As struDgv取芯和RQD) ss.tcbh = tcbh) Then
                        lo = publst取芯率RQD设置表.Find(Function(ss As struDgv取芯和RQD) ss.tcbh = tcbh).qxllo
                        up = publst取芯率RQD设置表.Find(Function(ss As struDgv取芯和RQD) ss.tcbh = tcbh).qxlup
                        If lo = 0 Or up = 0 Then
                            qxl = 0
                        Else
                            qxl = getRndN(lo, up)
                        End If

                        lo = publst取芯率RQD设置表.Find(Function(ss As struDgv取芯和RQD) ss.tcbh = tcbh).RQDlo
                        up = publst取芯率RQD设置表.Find(Function(ss As struDgv取芯和RQD) ss.tcbh = tcbh).RQDup
                        If lo = 0 Or up = 0 Then
                            rqd = 0
                        Else
                            rqd = getRndN(lo, up)
                        End If

                        Dim t As New stru取芯RQD(dsd, Int(qxl), Int(rqd))
                        lstqxlrqd.Add(t)
                    Else

                    End If
                    '此处测试间距变为随机值
                    dsd += interval
                Else
                    Exit Do
                End If
            Loop
            Return lstqxlrqd
        Catch ex As Exception

        End Try
    End Function
    Public Function funcGetZkBosuLst(zktclst As List(Of 地层数据结构)) As List(Of struBosu)
        Dim lstbosu As New List(Of struBosu)
        lstbosu.Clear()
        Try


            Dim dsd As Double
            Dim interval As Double
            Try
                interval = tbx波速测试间距.Text
                If interval <= 0 Then
                    MsgBox("波速起始测试深度和间距不能小于等于0")
                    Return lstbosu
                End If
            Catch ex As Exception
                MsgBox("波速起始测试深度和间距必须为数字")
                Return lstbosu
            End Try
            Try
                dsd = tbx波速起始测试深度.Text
                If dsd <= 0 Then
                    MsgBox("波速起始测试深度和间距不能小于等于0")
                    Return lstbosu
                End If
            Catch ex As Exception
                MsgBox("波速起始测试深度和间距必须为数字")
                Return lstbosu
            End Try
            Do While True

                Dim tc As 地层数据结构
                Dim bosuheng As Double
                Dim bosuzong As Double
                If zktclst.Exists(Function(ss As 地层数据结构) ss.cengdingshendu < dsd And ss.cengdishengdu >= dsd) Then
                    tc = zktclst.Find(Function(ss As 地层数据结构) ss.cengdingshendu < dsd And ss.cengdishengdu >= dsd)
                    Dim tcbh As String
                    tcbh = tc.TCZCBH & "-" & tc.TCYCBH
                    Dim lo As Double
                    Dim up As Double
                    If publst波速表土层设置.Exists(Function(ss As struDgv波速表设置) ss.tcbh = tcbh) Then
                        lo = publst波速表土层设置.Find(Function(ss As struDgv波速表设置) ss.tcbh = tcbh).loLmt横波
                        up = publst波速表土层设置.Find(Function(ss As struDgv波速表设置) ss.tcbh = tcbh).UpLmt横波
                        bosuheng = getRndN(lo, up)
                        lo = publst波速表土层设置.Find(Function(ss As struDgv波速表设置) ss.tcbh = tcbh).loLmt纵波
                        up = publst波速表土层设置.Find(Function(ss As struDgv波速表设置) ss.tcbh = tcbh).UpLmt纵波
                        bosuzong = getRndN(lo, up)
                        Dim t As New struBosu(dsd, Int(bosuheng), Int(bosuzong))
                        lstbosu.Add(t)
                    Else

                    End If
                    '此处测试间距变为随机值
                    dsd += interval
                Else
                    Exit Do
                End If
            Loop
            Return lstbosu
        Catch ex As Exception
            MsgBox("生成波速集合出错" & Err.Description)
            Return lstbosu
        End Try
    End Function
    Function dt() As Boolean


    End Function


    Public Structure 岩土线性指标值
        Dim 指标 As Double
        Dim 值 As Double
    End Structure
    Public Enum 文本文件中测试指标排列顺序枚举
        递增 = 0
        递减 = 1
    End Enum
    Function 从二维关系表中取值(二维表集合 As List(Of String), 第一指标 As Double, 第二指标 As Double) As Double
        Dim 内部二维表集合 As New List(Of String)
        For i As Integer = 0 To 二维表集合.Count - 1
            内部二维表集合.Add(二维表集合(i))
        Next
        '内部二维表集合 = 二维表集合
        Dim 第一行文本数据的集合 As New List(Of String)
        Dim s() As String = Split(二维表集合(0), vbTab)
        For i As Integer = 0 To UBound(s)
            第一行文本数据的集合.Add(s(i))
        Next

        内部二维表集合.RemoveAt(0)
        For iii As Integer = 0 To 内部二维表集合.Count - 1
            'MsgBox(内部二维表集合(iii))
        Next
        '定义当第一指标和第二指标超出范围的处理办法
        If 第一指标 < Split(内部二维表集合(0), vbTab)(0) Or 第一指标 > Split(内部二维表集合(内部二维表集合.Count - 1), vbTab)(0) Then Return -1
        '根据第一指标找相近的一行
        Dim 选定行文本数据的集合 As New List(Of String）
        Dim ss() As String = Split(内部二维表集合.Find(Function(ssss As String) Int(第一指标) - CDbl(Split(ssss, vbTab)(0)) <= 0), vbTab)
        For i As Integer = 0 To UBound(ss)
            选定行文本数据的集合.Add(ss(i))
        Next


        Dim l As New List(Of 岩土线性指标值）
        For i As Integer = 1 To 选定行文本数据的集合.Count - 1
            Dim ii As 岩土线性指标值
            If 选定行文本数据的集合(i) <> "" Then
                ii.指标 = 第一行文本数据的集合(i)

                ii.值 = 选定行文本数据的集合(i)
                l.Add(ii)
            End If
        Next
        Return 从集合中取得内插值(第二指标, l, 文本文件中测试指标排列顺序枚举.递增)
    End Function
    Function 从集合中取得内插值(测试指标 As Double, l As List(Of 岩土线性指标值)， 指标排列顺序 As 文本文件中测试指标排列顺序枚举) As Double
        If 测试指标 < l(0).指标 Or 测试指标 > l(l.Count - 1).指标 Then Return -1
        Select Case 指标排列顺序
            Case 文本文件中测试指标排列顺序枚举.递增
                If 测试指标 <= l(0).指标 Then
                    Return l(0).值
                ElseIf 测试指标 >= l(l.Count - 1).指标 Then
                    Return l(l.Count - 1).值
                Else
                    Dim ii As Integer = l.FindIndex(Function(s As 岩土线性指标值) 测试指标 <= s.指标) - 1
                    Return l(ii).值 + (l(ii + 1).值 - l(ii).值) * (测试指标 - l(ii).指标) / (l(ii + 1).指标 - l(ii).指标)
                End If
            Case 文本文件中测试指标排列顺序枚举.递减
                If 测试指标 >= l(0).指标 Then
                    Return l(0).值
                ElseIf 测试指标 <= l(l.Count - 1).指标 Then
                    Return l(l.Count - 1).值
                Else
                    Dim ii As Integer = l.FindIndex(Function(s As 岩土线性指标值) 测试指标 >= s.指标) - 1
                    Return l(ii).值 + (l(ii + 1).值 - l(ii).值) * (测试指标 - l(ii).指标) / (l(ii + 1).指标 - l(ii).指标)

                End If
        End Select


    End Function
    Sub qy()
        Try
            '取样
            ReDim qyRelatedZKs(0)
            Dim dtbl(dgvquyang.RowCount - 1) As qyywcsOptor.typeDgvTbl
            '读控件
            For i As Integer = 0 To dtbl.Length - 1
                With dtbl(i)
                    .tcbh = dgvquyang.Item(0, i).Value
                    .tcmc = dgvquyang.Item(1, i).Value



                    If IsNumeric(dgvquyang.Item(5, i).Value) Then
                        If dgvquyang.Item(5, i).Value <= 0 Then
                            .qycd = 0.2
                        Else
                            .qycd = dgvquyang.Item(5, i).Value
                        End If
                    End If

                    If IsNumeric(Trim(dgvquyang.Item(9, i).Value)) Then .MinLenght = dgvquyang.Item(9, i).Value
                    If IsNumeric(dgvquyang.Item(10, i).Value) Then
                        If dgvquyang.Item(10, i).Value > 0 Then
                            .qyCount = Int(dgvquyang.Item(10, i).Value)
                        End If

                    End If

                End With
            Next

            '写入数据库

            Dim qyMinsd As Double
            If IsNumeric(tbxqyMinsd.Text) = False Then
                MsgBox("最小取样深度不是一个数字,最小取样深度按0考虑")
            ElseIf tbxqyMinsd.Text < 0 Then
                MsgBox("最小取样深度小于0,最小取样深度按0考虑")
            Else
                qyMinsd = tbxqyMinsd.Text

            End If
            Dim qySpace As Double
            If IsNumeric(txQySpace.Text) = False Then
                MsgBox("取样间距不是一个数字,取样间距按1.8m考虑")
            ElseIf txQySpace.Text < 0 Then
                MsgBox("取样间距小于0,取样间距按1.8m考虑")
            Else
                qySpace = txQySpace.Text

            End If
            Dim zks As Object
            zks = Split(txtbxqyspecifiedZks.Text, ",", -1, CompareMethod.Text)
            WriteQyOfZks(zks, qyMinsd, qySpace, dtbl, qyywcsOptor.QyETcType.soil)
            zks = Split(txtbxqyspecifiedZks1.Text, ",", -1, CompareMethod.Text)
            WriteQyOfZks(zks, qyMinsd, qySpace, dtbl, qyywcsOptor.QyETcType.rock)

            'MsgBox("取样数据已入库")

        Catch ex As Exception
            MsgBox("取样入库过程错误-" & Err.Description)
        End Try

    End Sub
    Sub WriteQyOfZks(ByRef zks As Object, ByVal qyMinsd As Double, ByVal qySpace As Double, ByRef dtbl() As qyywcsOptor.typeDgvTbl, ByVal t As qyywcsOptor.QyETcType)
        If UBound(zks) >= 0 Then
            For n As Integer = 0 To UBound(zks)
                zks(n) = Trim(zks(n))
            Next
        Else

        End If
        If UBound(zks) >= 0 Then
            If UBound(zks) = 0 And zks(0) = "" Then Exit Sub

            For n As Integer = 0 To UBound(zks)
                lblStatus.Text = "正在对" & zks(n) & "添加取样..."
                lblStatus.Refresh()
                Dim rstc As New ADODB.Recordset
                '先吧钻孔土层数据选出
                rstc.Open("select * from z_g_tuceng where zkbh='" & Trim(zks(n)) & "' order by tccdsd", cnt, CursorTypeEnum.adOpenStatic)
                If rstc.RecordCount = 0 Then
                    'MsgBox("数据库中未找到" & zks(n) & "的土层数据")
                    rstc.Close()
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rstc)

                    GoTo 10
                End If
                '获取tc
                Dim tc(rstc.RecordCount - 1) As qyywcsOptor.typetc
                rstc.MoveFirst()
                Dim cdsd As Double
                For m As Integer = 0 To tc.Length - 1
                    If IsDBNull(rstc.Fields("gcsy").Value) = False Then tc(m).gcsy = rstc.Fields("gcsy").Value
                    tc(m).tcbh = rstc.Fields("tczcbh").Value & "-" & rstc.Fields("tcycbh").Value
                    If rstc.Fields("tcmc").Value Like "*岩" Then
                        tc(m).tcType = qyywcsOptor.QyETcType.rock
                    Else
                        tc(m).tcType = qyywcsOptor.QyETcType.soil
                    End If
                    If m = 0 Then
                        tc(m).tclen = rstc.Fields("tccdsd").Value
                    Else
                        tc(m).tclen = rstc.Fields("tccdsd").Value - cdsd
                    End If
                    tc(m).tcdsd = rstc.Fields("tccdsd").Value
                    tc(m).tcdingSD = rstc.Fields("tccdsd").Value - tc(m).tclen
                    cdsd = rstc.Fields("tccdsd").Value
                    rstc.MoveNext()
                Next

                rstc.Close()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rstc)

                Dim opt As New qyywcsOptor
                Dim zkqyarr() As qyywcsOptor.TypeQuyang
                ReDim zkqyarr(0)


                opt.getZKQyArr(zkqyarr, tc, dtbl, qyMinsd, qySpace, t)
                '如果没有生成zkqyarr则进入下一钻孔
                If zkqyarr.Length >= 1 And zkqyarr(0).qydsd > 0 Then
                    'If Not (qyRelatedZKs.Length = 1 And qyRelatedZKs(0).zkbh = "") Then ReDim Preserve qyRelatedZKs(bgRelatedZKs.Length)
                    'qyRelatedZKs(qyRelatedZKs.Length - 1).zkbh = Trim(zks(n))
                Else
                    GoTo 10
                End If

                '写入数据库
                Dim rsqy As New ADODB.Recordset
                rsqy.Open("select * from z_c_QuYang where zkbh='" & Trim(zks(n)) & "'", cnt, CursorTypeEnum.adOpenStatic)
                '先看看本钻孔中是否有取样
                If rsqy.RecordCount > 0 Then
                    '如果有
                    If cbxoverwriteqy.Checked Then
                        Dim cmd As New ADODB.Command
                        cmd.ActiveConnection = cnt
                        cmd.CommandText = "delete * from z_c_quyang where zkbh='" & Trim(zks(n)) & "'"
                        cmd.Execute()
                        For ii As Integer = 0 To zkqyarr.Length - 1
                            cmd.CommandText = "insert into z_c_quyang (zkbh,gcsy,qybh,qysd,qyhd,qylx) values('" & Trim(zks(n)) & "'," & tc(0).gcsy & "," & zkqyarr(ii).qybh & "," & zkqyarr(ii).qydsd & "," & zkqyarr(ii).qycd & ",0)"
                            cmd.Execute()
                        Next
                    End If
                Else
                    '如果没有
                    Dim cmd As New ADODB.Command
                    cmd.ActiveConnection = cnt
                    For ii As Integer = 0 To zkqyarr.Length - 1
                        cmd.CommandText = "insert into z_c_quyang (zkbh,gcsy,qybh,qysd,qyhd,qylx) values('" & Trim(zks(n)) & "'," & tc(0).gcsy & "," & zkqyarr(ii).qybh & "," & zkqyarr(ii).qydsd & "," & zkqyarr(ii).qycd & ",0)"
                        cmd.Execute()
                    Next
                End If
                rsqy.Close()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rsqy)
10:         Next
        End If
    End Sub
    Sub bg()
        Try
            '标贯
            ReDim bgRelatedZKs(0)
            Dim dtbl(dgvbg.RowCount - 1) As qyywcsOptor.typeDgvTbl
            '读控件
            For i As Integer = 0 To dtbl.Length - 1
                With dtbl(i)
                    .tcbh = dgvbg.Item(0, i).Value
                    .tcmc = dgvbg.Item(1, i).Value
                    If dgvbg.Item(2, i).Value = "N10" Then
                        .cslx = qyywcsOptor.eDtlx.N10
                    ElseIf dgvbg.Item(2, i).Value = "N635" Then
                        .cslx = qyywcsOptor.eDtlx.N635
                    ElseIf dgvbg.Item(2, i).Value = "N120" Then
                        .cslx = qyywcsOptor.eDtlx.N120
                    Else
                        .cslx = qyywcsOptor.eDtlx.N10
                    End If



                    If UBound(Split(dgvbg.Item(5, i).Value, "-", -1, CompareMethod.Text)) >= 1 Then
                        If IsNumeric(Trim(Split(dgvbg.Item(5, i).Value, "-", -1, CompareMethod.Text)(0))) And IsNumeric(Trim(Split(dgvbg.Item(5, i).Value, "-", -1, CompareMethod.Text)(1))) Then
                            .loLmt = Val(Trim(Split(dgvbg.Item(5, i).Value, "-", -1, CompareMethod.Text)(0)))
                            .UpLmt = Val(Trim(Split(dgvbg.Item(5, i).Value, "-", -1, CompareMethod.Text)(1)))
                        End If

                    End If

                    If IsNumeric(Trim(dgvbg.Item(9, i).Value)) Then .MinLenght = dgvbg.Item(9, i).Value
                    If IsNumeric(Trim(dgvbg.Item(10, i).Value)) Then .testCount = Int(dgvbg.Item(10, i).Value)
                End With
            Next
            '写入数据库
            Dim zks As Object = Split(txtbxbgspecifiedZks.Text, ",", -1, CompareMethod.Text)
            If UBound(zks) >= 0 Then
                For n As Integer = 0 To UBound(zks)
                    zks(n) = Trim(zks(n))
                Next
            Else

            End If
            If UBound(zks) >= 0 Then
                If UBound(zks) = 0 And zks(0) = "" Then Exit Sub

                For n As Integer = 0 To UBound(zks)
                    lblStatus.Text = "正在对" & zks(n) & "添加标贯..."
                    lblStatus.Refresh()
                    Dim rstc As New ADODB.Recordset
                    '先吧钻孔土层数据选出
                    rstc.Open("select * from z_g_tuceng where zkbh='" & Trim(zks(n)) & "' order by tccdsd", cnt, CursorTypeEnum.adOpenStatic)
                    If rstc.RecordCount = 0 Then
                        'MsgBox("数据库中未找到" & zks(n) & "的土层数据")
                        rstc.Close()
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(rstc)

                        GoTo 10
                    End If
                    '获取tc
                    Dim tc(rstc.RecordCount - 1) As qyywcsOptor.typetc
                    rstc.MoveFirst()
                    Dim cdsd As Double
                    For m As Integer = 0 To tc.Length - 1
                        If IsDBNull(rstc.Fields("gcsy").Value) = False Then tc(m).gcsy = rstc.Fields("gcsy").Value
                        tc(m).tcbh = rstc.Fields("tczcbh").Value & "-" & rstc.Fields("tcycbh").Value

                        If m = 0 Then
                            tc(m).tclen = rstc.Fields("tccdsd").Value
                        Else
                            tc(m).tclen = rstc.Fields("tccdsd").Value - cdsd
                        End If
                        tc(m).tcdsd = rstc.Fields("tccdsd").Value
                        tc(m).tcdingSD = rstc.Fields("tccdsd").Value - tc(m).tclen
                        cdsd = rstc.Fields("tccdsd").Value
                        rstc.MoveNext()
                    Next

                    rstc.Close()
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rstc)

                    Dim opt As New qyywcsOptor
                    Dim zkBGarr() As qyywcsOptor.TypeBG
                    ReDim zkBGarr(0)
                    Dim bgMinsd As Double
                    If IsNumeric(tbxbgminsd.Text) = False Then
                        MsgBox("最小标贯深度不是一个数字,最小标贯深度按0考虑")
                    ElseIf tbxbgminsd.Text < 0 Then
                        MsgBox("最小标贯深度小于0,最小标贯深度按0考虑")
                    Else
                        bgMinsd = tbxbgminsd.Text

                    End If
                    Dim bgSpace As Double
                    If IsNumeric(txBgSpace.Text) = False Then
                        MsgBox("标贯间距不是一个数字,标贯间距按1.8m考虑")
                    ElseIf txBgSpace.Text < 0 Then
                        MsgBox("标贯间距小于0,标贯间距按1.8m考虑")
                    Else
                        bgSpace = txBgSpace.Text

                    End If

                    opt.getZkBgArr(zkBGarr, tc, dtbl, bgMinsd, bgSpace)

                    If zkBGarr.Length >= 1 And zkBGarr(0).bggc > 0 Then
                        If Not (bgRelatedZKs.Length = 1 And bgRelatedZKs(0).zkbh = "") Then ReDim Preserve bgRelatedZKs(bgRelatedZKs.Length)
                        bgRelatedZKs(bgRelatedZKs.Length - 1).zkbh = Trim(zks(n))
                    Else
                        GoTo 10
                    End If

                    '写入数据库
                    Dim rsbg As New ADODB.Recordset
                    rsbg.Open("select * from z_y_biaoguan where zkbh='" & Trim(zks(n)) & "'", cnt, CursorTypeEnum.adOpenStatic)
                    If rsbg.RecordCount > 0 Then
                        If cbxoverwritebg.Checked = True Then
                            Dim cmd As New ADODB.Command
                            cmd.ActiveConnection = cnt
                            cmd.CommandText = "delete * from z_y_biaoguan where zkbh='" & Trim(zks(n)) & "'"
                            cmd.Execute()
                            For ii As Integer = 0 To zkBGarr.Length - 1
                                cmd.CommandText = "insert into z_y_biaoguan (zkbh,gcsy,bgdsd,bgjs,bglx,bggc,cy,bgtzz,bgxjs_,bgxjstj_,bgjs_,bgjstj_,bgyzcd) values('" & Trim(zks(n)) & "'," & tc(0).gcsy & "," & zkBGarr(ii).dsd & "," & zkBGarr(ii).bgjs & "," & 1 & "," & zkBGarr(ii).bggc & ",1,0,1,1,1,1,0.3)"
                                cmd.Execute()
                            Next
                        End If
                    Else
                        Dim cmd As New ADODB.Command
                        cmd.ActiveConnection = cnt
                        For ii As Integer = 0 To zkBGarr.Length - 1
                            cmd.CommandText = "insert into z_y_biaoguan (zkbh,gcsy,bgdsd,bgjs,bglx,bggc,cy,bgtzz,bgxjs_,bgxjstj_,bgjs_,bgjstj_,bgyzcd) values('" & Trim(zks(n)) & "'," & tc(0).gcsy & "," & zkBGarr(ii).dsd & "," & zkBGarr(ii).bgjs & "," & 1 & "," & zkBGarr(ii).bggc & ",1,0,1,1,1,1,0.3)"
                            cmd.Execute()
                        Next
                    End If
                    rsbg.Close()
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rsbg)
10:             Next
            End If

            'MsgBox("标贯数据已入库")

        Catch ex As Exception
            MsgBox("标贯入库过程错误-" & Err.Description)
        End Try

    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        dt()
        bg()
        qy()
        生成波速取芯率和RQD()
        lblStatus.Text = "操作完成"
        lblStatus.Refresh()
    End Sub
    Structure qyarr
        Dim tcbh As String
        Dim qycount As Integer
    End Structure
    Sub getSummary()
        Try
            txtbxSummary.Clear()
            Dim rs As New Recordset
            Dim s As New qySelectTool
            'rs.Close()
            rs.Open("select * from z_zuankong ", cnt, CursorTypeEnum.adOpenStatic)
            Dim zkcount As Integer = rs.RecordCount
            rs.Close()
            rs.Open("select * from z_y_dongtan where dtjs>0 order by zkbh ", cnt, CursorTypeEnum.adOpenStatic)

            Dim dtcount As Integer
            If s.getzkbharr(rs)(0) <> "" Then
                dtcount = s.getzkbharr(rs).Length
            Else
                dtcount = 0
            End If
            rs.Close()
            rs.Open("select * from z_y_biaoguan order by zkbh", cnt, CursorTypeEnum.adOpenStatic)
            Dim bgcount As Integer
            If s.getzkbharr(rs)(0) <> "" Then
                bgcount = s.getzkbharr(rs).Length
            Else
                bgcount = 0
            End If

            rs.Close()
            rs.Open("select * from z_c_quyang order by zkbh", cnt, CursorTypeEnum.adOpenStatic)
            Dim qycount As Integer
            If s.getzkbharr(rs)(0) <> "" Then
                qycount = s.getzkbharr(rs).Length
            Else
                qycount = 0
            End If

            rs.Close()
            txtbxSummary.Text = "数据库中共有钻孔数：" & zkcount & vbCrLf & "其中：" & vbCrLf & "共有动探孔数：" & dtcount & vbCrLf & "共有标贯孔数：" & bgcount & vbCrLf & "共有取样孔数：" & qycount & vbCrLf & "其中:" & vbCrLf
            rs.Open("SELECT z_c_quyang.zkbh,z_c_quyang.qybh,z_c_quyang.qysd,z_c_quyang.qyhd,z_g_tuceng.zkbh,z_g_tuceng.tccdsd,z_g_tuceng.tchd,z_g_tuceng.tczcbh,z_g_tuceng.tcycbh ,z_g_tuceng.tcmc,z_g_tuceng.tcfhcd,z_g_tuceng.tcmsd,z_g_tuceng.tcksx,z_g_tuceng.tcsid from z_c_quyang , z_g_tuceng where z_c_quyang.zkbh=z_g_tuceng.zkbh and (z_c_quyang.qysd<z_g_tuceng.tccdsd and z_c_quyang.qysd>z_g_tuceng.tccdsd-z_g_tuceng.tchd) order by z_g_tuceng.tczcbh,z_g_tuceng.tcycbh", cnt, CursorTypeEnum.adOpenStatic)
            If rs.RecordCount = 0 Then
                If qycount > 0 Then
                    MsgBox("无法生成取样表，可能该数据库未经理正数检合格。请先在理正中对数据库进行数检")
                Else
                    MsgBox("取样数为0，无法生成取样表。")
                End If

                Exit Sub
            End If
            rs.MoveFirst()
            Dim q() As qyarr
            ReDim q(0)
            Do Until rs.EOF
                If q(0).tcbh = "" Then
                    With q(0)
                        .tcbh = rs.Fields("tczcbh").Value & "-" & rs.Fields("tcycbh").Value
                        .qycount = .qycount + 1
                    End With

                ElseIf q(q.Length - 1).tcbh = rs.Fields("tczcbh").Value & "-" & rs.Fields("tcycbh").Value Then
                    q(q.Length - 1).qycount = q(q.Length - 1).qycount + 1
                Else
                    ReDim Preserve q(q.Length)
                    With q(q.Length - 1)
                        .tcbh = rs.Fields("tczcbh").Value & "-" & rs.Fields("tcycbh").Value
                        .qycount = .qycount + 1
                    End With
                End If
                rs.MoveNext()
            Loop


            For i As Integer = 0 To q.Length - 1
                txtbxSummary.Text = txtbxSummary.Text & q(i).tcbh & "取样数：" & q(i).qycount & vbCrLf
            Next
            txtbxSummary.Text = txtbxSummary.Text & "取样单：" & vbCrLf
            rs.MoveFirst()
            txtbxSummary.Text = txtbxSummary.Text & "钻孔编号" & vbTab & "岩土名称" & vbTab & "取样深度" & vbTab & "取样长度" & vbTab & "湿度" & vbTab & "密实度" & vbTab & "可塑性" & vbTab & "风化程度" & vbCrLf
            Do Until rs.EOF
                txtbxSummary.Text = txtbxSummary.Text & rs.Fields("z_c_quyang.zkbh").Value & vbTab & rs.Fields("tcmc").Value & vbTab & rs.Fields("qysd").Value & vbTab & rs.Fields("qyhd").Value & vbTab & rs.Fields("tcsid").Value & vbTab & rs.Fields("tcmsd").Value & vbTab & rs.Fields("tcksx").Value & vbTab & rs.Fields("tcfhcd").Value & vbCrLf
                rs.MoveNext()

            Loop
            rs.Close()
        Catch ex As Exception
            MsgBox(Err.Description)
        End Try



    End Sub
    Private Sub dgvywcs_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvDt.CellContentClick
        Try
            If dgvDt.CurrentCell.ColumnIndex = 10 Then
                Dim rs As New ADODB.Recordset
                rs.Open("select * from z_g_tuceng where tczcbh=" & Split(dgvDt.Item(0, dgvDt.CurrentCell.RowIndex).Value)(0) & " and tcycbh=" & Split(dgvDt.Item(0, dgvDt.CurrentCell.RowIndex).Value)(1) & " order by tczkbh")
                rs.MoveFirst()

                Do Until rs.EOF

                Loop
            End If
        Catch ex As Exception

        End Try

    End Sub





    Private Sub rbtselectallzks_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtdtselectallzks.CheckedChanged
        Try
            txtbxdtspecifiedZks.Clear()
            If rbtdtselectallzks.Checked Then
                Dim rs As New Recordset
                rs.Open("select * from z_zuankong order by zkbh", cnt, CursorTypeEnum.adOpenStatic)
                If rs.RecordCount > 0 Then
                    Dim z As New qySelectTool
                    Dim o As Object = z.getzkbharr(rs)

                    For i As Integer = 0 To UBound(o)
                        If i < UBound(o) Then
                            txtbxdtspecifiedZks.Text = txtbxdtspecifiedZks.Text & o(i) & ","
                        Else
                            txtbxdtspecifiedZks.Text = txtbxdtspecifiedZks.Text & o(i)
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(Err.Description)
        End Try
    End Sub

    Private Sub rbtdtselectoddzks_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtdtselectoddzks.CheckedChanged
        Try
            txtbxdtspecifiedZks.Clear()
            If rbtdtselectoddzks.Checked Then
                Dim rs As New Recordset
                rs.Open("select * from z_zuankong order by zkbh", cnt, CursorTypeEnum.adOpenStatic)
                If rs.RecordCount > 0 Then
                    Dim z As New qySelectTool
                    Dim o As Object = z.getzkbharr(rs)

                    For i As Integer = 0 To UBound(o) Step 2
                        If i < UBound(o) Then
                            txtbxdtspecifiedZks.Text = txtbxdtspecifiedZks.Text & o(i) & ","
                        Else
                            txtbxdtspecifiedZks.Text = txtbxdtspecifiedZks.Text & o(i)
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(Err.Description)
        End Try
    End Sub

    Private Sub rbtdtselectevenzks_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtdtselectevenzks.CheckedChanged
        Try
            txtbxdtspecifiedZks.Clear()
            If rbtdtselectevenzks.Checked Then
                Dim rs As New Recordset
                rs.Open("select * from z_zuankong order by zkbh", cnt, CursorTypeEnum.adOpenStatic)
                If rs.RecordCount > 0 Then
                    Dim z As New qySelectTool
                    Dim o As Object = z.getzkbharr(rs)
                    If UBound(o) >= 1 Then
                        For i As Integer = 1 To UBound(o) Step 2
                            If i < UBound(o) Then
                                txtbxdtspecifiedZks.Text = txtbxdtspecifiedZks.Text & o(i) & ","
                            Else
                                txtbxdtspecifiedZks.Text = txtbxdtspecifiedZks.Text & o(i)
                            End If
                        Next
                    End If

                End If
            End If
        Catch ex As Exception
            MsgBox(Err.Description)
        End Try
    End Sub
    Private Sub rbtBSselectevenzks_CheckedChanged(sender As Object, e As EventArgs) Handles rbtBSselectevenzks.CheckedChanged, rbtBSselectallzks.CheckedChanged, rbtBSselectoddzks.CheckedChanged
        Try

            Select Case CType(sender, RadioButton).Name
                Case rbtBSselectallzks.Name
                    txtbxbosuspecifiedZks.Clear()

                    If rbtBSselectallzks.Checked Then
                        txtbxbosuspecifiedZks.Text = getZkBh(QyESelectMode.selectAll)
                    End If

                Case rbtBSselectoddzks.Name
                    txtbxbosuspecifiedZks.Clear()
                    If rbtBSselectoddzks.Checked Then
                        txtbxbosuspecifiedZks.Text = getZkBh(QyESelectMode.selectOdd)
                    End If
                Case rbtBSselectevenzks.Name
                    txtbxbosuspecifiedZks.Clear()
                    If rbtBSselectevenzks.Checked Then
                        txtbxbosuspecifiedZks.Text = getZkBh(QyESelectMode.selectEven)
                    End If
            End Select

        Catch ex As Exception

        End Try
    End Sub

    Private Sub rbtqyselectallzks_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtqyselectoddzks.CheckedChanged, rbtqyselectevenzks.CheckedChanged, rbtqyselectallzks.CheckedChanged
        Try
            Select Case CType(sender, RadioButton).Name
                Case "rbtqyselectallzks"
                    Select Case rbntGeoTypesoil.Checked
                        Case True
                            txtbxqyspecifiedZks.Clear()
                            If rbtqyselectallzks.Checked Then
                                txtbxqyspecifiedZks.Text = getZkBh(QyESelectMode.selectAll)
                            End If
                        Case False
                            txtbxqyspecifiedZks1.Clear()
                            If rbtqyselectallzks.Checked Then
                                txtbxqyspecifiedZks1.Text = getZkBh(QyESelectMode.selectAll)
                            End If
                    End Select
                Case "rbtqyselectoddzks"
                    Select Case rbntGeoTypesoil.Checked
                        Case True
                            txtbxqyspecifiedZks.Clear()
                            If rbtqyselectoddzks.Checked Then
                                txtbxqyspecifiedZks.Text = getZkBh(QyESelectMode.selectOdd)
                            End If
                        Case False
                            txtbxqyspecifiedZks1.Clear()
                            If rbtqyselectoddzks.Checked Then
                                txtbxqyspecifiedZks1.Text = getZkBh(QyESelectMode.selectOdd)
                            End If
                    End Select
                Case "rbtqyselectevenzks"
                    Select Case rbntGeoTypesoil.Checked
                        Case True
                            txtbxqyspecifiedZks.Clear()
                            If rbtqyselectevenzks.Checked Then
                                txtbxqyspecifiedZks.Text = getZkBh(QyESelectMode.selectEven)
                            End If
                        Case False
                            txtbxqyspecifiedZks1.Clear()
                            If rbtqyselectevenzks.Checked Then
                                txtbxqyspecifiedZks1.Text = getZkBh(QyESelectMode.selectEven)
                            End If
                    End Select
            End Select


        Catch ex As Exception
            MsgBox(Err.Description)
        End Try
    End Sub
    Enum QyESelectMode
        selectAll = 1

        selectOdd = 2
        selectEven = 3
    End Enum
    Function getZkBh(ByVal selectMode As QyESelectMode) As String
        Dim t As String
        'Dim rs As New Recordset
        'rs.Open("select * from z_zuankong order by zkbh", cnt, CursorTypeEnum.adOpenStatic)
        Dim z As New qySelectTool
        Dim o() As String  '= z.getzkbharr(rs)
        o = PubList本工程钻孔表集合.Select(Of String)(Function(ss As 钻孔编号高程数据结构) ss.zkbh).ToArray
        Select Case selectMode
            Case QyESelectMode.selectAll
                For i As Integer = 0 To UBound(o)
                    If i < UBound(o) Then
                        t = t & o(i) & ","
                    Else
                        t = t & o(i)
                    End If
                Next
            Case QyESelectMode.selectEven
                If UBound(o) >= 1 Then
                    For i As Integer = 1 To UBound(o) Step 2
                        If i < UBound(o) Then
                            t = t & o(i) & ","
                        Else
                            t = t & o(i)
                        End If
                    Next
                End If
            Case QyESelectMode.selectOdd
                For i As Integer = 0 To UBound(o) Step 2
                    If i < UBound(o) Then
                        t = t & o(i) & ","
                    Else
                        t = t & o(i)
                    End If
                Next
        End Select
        Return t

    End Function



    Private Sub rbtbgselectallzks_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtbgselectallzks.CheckedChanged
        Try
            txtbxbgspecifiedZks.Clear()
            If rbtbgselectallzks.Checked Then
                Dim rs As New Recordset
                rs.Open("select * from z_zuankong order by zkbh", cnt, CursorTypeEnum.adOpenStatic)
                If rs.RecordCount > 0 Then
                    Dim z As New qySelectTool
                    Dim o As Object = z.getzkbharr(rs)

                    For i As Integer = 0 To UBound(o)
                        If i < UBound(o) Then
                            txtbxbgspecifiedZks.Text = txtbxbgspecifiedZks.Text & o(i) & ","
                        Else
                            txtbxbgspecifiedZks.Text = txtbxbgspecifiedZks.Text & o(i)
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(Err.Description)
        End Try
    End Sub

    Private Sub rbtbgselectoddzks_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtbgselectoddzks.CheckedChanged
        Try
            txtbxbgspecifiedZks.Clear()
            If rbtbgselectoddzks.Checked Then
                Dim rs As New Recordset
                rs.Open("select * from z_zuankong order by zkbh", cnt, CursorTypeEnum.adOpenStatic)
                If rs.RecordCount > 0 Then
                    Dim z As New qySelectTool
                    Dim o As Object = z.getzkbharr(rs)
                    For i As Integer = 0 To UBound(o) Step 2
                        If i < UBound(o) Then
                            txtbxbgspecifiedZks.Text = txtbxbgspecifiedZks.Text & o(i) & ","
                        Else
                            txtbxbgspecifiedZks.Text = txtbxbgspecifiedZks.Text & o(i)
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(Err.Description)
        End Try
    End Sub

    Private Sub rbtbgselectevenzks_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtbgselectevenzks.CheckedChanged
        Try
            txtbxbgspecifiedZks.Clear()
            If rbtbgselectevenzks.Checked Then
                Dim rs As New Recordset
                rs.Open("select * from z_zuankong order by zkbh", cnt, CursorTypeEnum.adOpenStatic)
                If rs.RecordCount > 0 Then
                    Dim z As New qySelectTool
                    Dim o As Object = z.getzkbharr(rs)
                    If UBound(o) >= 1 Then
                        For i As Integer = 1 To UBound(o) Step 2
                            If i < UBound(o) Then
                                txtbxbgspecifiedZks.Text = txtbxbgspecifiedZks.Text & o(i) & ","
                            Else
                                txtbxbgspecifiedZks.Text = txtbxbgspecifiedZks.Text & o(i)
                            End If
                        Next
                    End If

                End If
            End If
        Catch ex As Exception
            MsgBox(Err.Description)
        End Try
    End Sub


    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntbiangenExcute.Click

    End Sub



    Private Sub GroupBox0_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        getSummary()
    End Sub

    Private Sub Form1_Closed(sender As Object, e As EventArgs) Handles Me.Closed

    End Sub

    Private Sub GroupBox8_Enter(sender As Object, e As EventArgs) Handles GroupBox8.Enter

    End Sub

    Private Sub Button3_Click_2(sender As Object, e As EventArgs) Handles Button3.Click
        Dim cmd As New ADODB.Command
        cmd.ActiveConnection = cnt

        cmd.CommandText = "delete * from z_y_bosu"
        Try
            cmd.Execute()
        Catch ex As Exception
            MsgBox("删除波速记录出错")

        End Try

    End Sub

    ''' <summary>
    ''' 生成word文档
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        '
        Me.ProgressBar1.Maximum = 100
        Me.ProgressBar1.Minimum = 0
        Me.Label21.Text = $"开始读取并导出钻孔信息!"
        If Me.OpenFileDialog1.FileName = String.Empty Then
            MsgBox（"请先读取数据再执行操做"）
            Call OpenDataBase_Click(OpenDataBase, Nothing)
            Exit Sub
        End If
        Dim dirApp = Path.GetDirectoryName(Reflection.Assembly.GetExecutingAssembly().Location) + "\用到的文件"

        Dim rowDis As Integer = Me.NumericUpDown1.Value
        Dim wordSample As String = "XY-100钻探记录.doc"

        Dim wordSampleFullPath As String = Path.Combine(dirApp, wordSample)

        If FileInUse(wordSampleFullPath) Then
            MsgBox(wordSampleFullPath + ",被占用或者被打开了,请关闭文档后再执行此操作!!!")
            Exit Sub
        End If

        Dim wordApp As Microsoft.Office.Interop.Word.Application = Nothing

        Dim FontName() As String = {"liguofu"} ' Dim FontName() As String = {"liguofu", "【嵐】芊柔体"}

        If 需求的字体的是否安装(FontName) Then
            If File.Exists(wordSampleFullPath) Then
                'Public PubList本工程钻孔表集合 As New List(Of 钻孔编号高程数据结构)
                'Public PubList本工程地层数据集合 As New List(Of 地层数据结构)
                If Me.PubList本工程钻孔表集合.Count > 0 And Me.PubList本工程地层数据集合.Count > 0 Then

                    Dim tcbhxx As New List(Of MyTCFill2WordData)
                    For Each item As 钻孔编号高程数据结构 In PubList本工程钻孔表集合
                        Dim temp地层数据 As List(Of 地层数据结构） = Me.PubList本工程地层数据集合.Where(Function(t As 地层数据结构) t.zkbh = item.zkbh).ToList()
                        Dim temp As MyTCFill2WordData = New MyTCFill2WordData(Me.tbx_工程名称.Text, Me.tbx_勘察单位.Text, item, temp地层数据)
                        tcbhxx.Add(temp)
                    Next
                    Dim wordAppType As Type = Type.GetTypeFromProgID("Word.Application")
                    wordApp = CType(Activator.CreateInstance(wordAppType), MSWord.Application)
                    wordApp.Visible = False
                    wordApp.DisplayAlerts = MSWord.WdAlertLevel.wdAlertsNone
                    Try
                        Dim progressVal As Integer = 1
                        Dim workLocation = Path.GetDirectoryName(Me.OpenFileDialog1.FileName)
                        For Each item As MyTCFill2WordData In tcbhxx
                            Me.Label21.Text = $"正在导出{item.ZKBH}钻孔信息,已经完成导出{progressVal}个钻孔信息,剩余{tcbhxx.Count - progressVal}个未导出,请等待!"
                            Me.Label21.Refresh()
                            Me.ProgressBar1.Value = Int((progressVal / tcbhxx.Count) * 100)
                            Me.ProgressBar1.Refresh()
                            Dim curDocName = Path.Combine(workLocation, $"{item.ZKBH}-钻探记录.doc")
                            If File.Exists(curDocName) Then File.Delete(curDocName)
                            WordHelper.SetTextBoxTextFont(wordApp, wordSampleFullPath, workLocation, FontName, rowDis, item)
                            progressVal = progressVal + 1
                        Next
                        wordApp.Quit()
                        killWordAppByProcess(wordApp)
                        MsgBox（"生成成功！"）
                    Catch ex As Exception
                        MsgBox（"生成失败！" & vbNewLine & ex.Message）
                        wordApp.Quit()
                        killWordAppByProcess(wordApp)
                    End Try
                End If
            Else
                If （MessageBox.Show($"模板文件缺失，请将模板放入到 {dirApp }路径下面,模板的名称是:{wordSample }", "打开路径查询模板", MessageBoxButtons.YesNo， MessageBoxIcon.Exclamation) = DialogResult.Yes) Then
                    System.Diagnostics.Process.Start("Explorer.exe", dirApp)
                End If
            End If
        Else
            If （MessageBox.Show($"字体未安装，请安装相应的字体{FontName(0)}后再重启本软件！", "打开字体安装包", MessageBoxButtons.YesNo， MessageBoxIcon.Exclamation) = DialogResult.Yes) Then
                'System.Diagnostics.Process.Start("Explorer.exe", dirApp)
                Dim processFontInstall As Process = New Process()
                Dim processFontInstall_1 As Process = New Process()
                With processFontInstall
                    .StartInfo.FileName = dirApp + "\" + FontName(0) + ".ttf"
                    .StartInfo.CreateNoWindow = False
                    .StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal
                    .Start()

                    'If processFontInstall.WaitForExit(2000) = True Then
                    '    With processFontInstall_1
                    '        .StartInfo.FileName = dirApp + "\" + FontName(1) + ".ttf"
                    '        .StartInfo.CreateNoWindow = False
                    '        .StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal
                    '        .Start()
                    '    End With
                    'End If
                End With


                Me.Close()
            End If
        End If


    End Sub

    Private Shared Sub killWordAppByProcess(wordApp As MSWord.Application)
        Dim myProcess = New Process()
        Dim wordProcess = Process.GetProcessesByName("winword")
        For Each pro As Process In wordProcess
            Dim str = pro.MainWindowTitle
            If str = "" Then
                pro.Kill()
            End If
        Next
    End Sub

    Function 需求的字体的是否安装(FontName() As String) As Boolean

        Dim Installfont As New System.Drawing.Text.InstalledFontCollection
        '开始枚举字体并加入到comFont中
        Dim sysfonts As New List(Of String)

        For Each ff As FontFamily In Installfont.Families
            sysfonts.Add(ff.Name)
        Next

        Dim requiredFontCount As Integer = 0

        For Each item As String In FontName
            If sysfonts.Contains(item) Then
                requiredFontCount = requiredFontCount + 1
            End If
        Next
        If requiredFontCount = FontName.Length Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub TabPage6_Click(sender As Object, e As EventArgs) Handles TabPage6.Click
        Me.NumericUpDown1.Visible = True
    End Sub

    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged

    End Sub

    Private Sub Label21_Click(sender As Object, e As EventArgs) Handles Label21.Click

    End Sub

    Function FileInUse(filename As String) As Boolean
        Dim fs As FileStream, use As Boolean = True
        Try
            fs = New FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.None)
            use = False
            fs.Close()
        Catch ex As Exception
            fs.Close()
        End Try
        Return use
    End Function
End Class

'Added BY Nans @20201108
Class MyTCFill2WordData
    Public GCMC As String '工程名称
    Public ZKBG As String '孔口标高
    Public ZKZJLX As String '钻机类型
    Public KCDW As String '勘察单位
    Public ZKBH As String '钻孔编号
    Public YR As String '年
    Public MO As String '月
    Public DA As String '日
    Public YR1 As String '年
    Public MO1 As String '月
    Public DA1 As String '日
    Public ListTCinfor As List(Of MyTCFill2WordDataTCinfo)
    Sub New(_gcmc As String, _kcdw As String, zkbhgcsj As 钻孔编号高程数据结构, 当前编号地层数据集合 As List(Of 地层数据结构))
        Me.GCMC = _gcmc
        Me.KCDW = _kcdw
        Me.ZKBG = zkbhgcsj.zkbg
        Me.ZKBH = zkbhgcsj.zkbh
        Me.ZKZJLX = zkbhgcsj.ZKZJLX
        Dim dateksrq As Date = CDate(Format(CInt(zkbhgcsj.ZKKSRQ), "0000-00-00"))
        Dim datezzrq As Date = CDate(Format(CInt(zkbhgcsj.ZKZZRQ), "0000-00-00"))

        Me.YR = dateksrq.Year
        Me.MO = dateksrq.Month
        Me.DA = dateksrq.Day

        Me.YR1 = datezzrq.Year
        Me.MO1 = datezzrq.Month
        Me.DA1 = datezzrq.Day

        Me.ListTCinfor = New List(Of MyTCFill2WordDataTCinfo)()
        For Each item As 地层数据结构 In 当前编号地层数据集合
            If item.zkbh = zkbhgcsj.zkbh Then
                Dim tcxx As MyTCFill2WordDataTCinfo
                With tcxx
                    .KSX = item.TCKSX
                    .YS = item.TCYS
                    .TCMC = item.tcmc
                    .MSD = item.TCMSD
                    .TCCDSD = item.cengdishengdu
                    .SD = item.TCSID
                    .MS = item.TCMS
                End With
                ListTCinfor.Add(tcxx)
            End If
        Next

    End Sub
End Class

Structure MyTCFill2WordDataTCinfo
    Public TCCDSD As String '进尺
    Public YS As String '颜色
    Public KSX As String '状态
    Public MSD As String '密实度
    Public SD As String '湿度
    Public MS As String '成   份   及   其  它
    Public TCMC As String '岩土名称
End Structure
