Imports System.Data
Imports System.Data.Common
Imports System.Collections.Generic
Imports System.Text
Imports ADODB
Imports System.Security.Cryptography
Imports System.IO
Public Structure Stru用户信息
    Dim username As String
    Dim 用户身份 As Stru用户身份
    'Dim freeCount As Integer
    Dim money As Double
    Dim 购买的产品 As String
    Dim validDate As String

End Structure
Public Enum Stru用户身份
    管理员 = 1
    试用用户 = 2
    正式用户 = 3
    不存在的用户 = 4
    过期用户或欠费用户 = 5
End Enum
Public Enum enumstatus用户状态
    未查询
    正常
    费用低
    欠费
    时限将近
    过期
    用户不存在
End Enum


Public Class Cls用户信息和费用余额查询扣减器
    ' Public rs As ADODB.Recordset
    'Public connbuilder As MySqlConnectionStringBuilder
    Public AdodbConn As ADODB.Connection
    Public OleDbCon As OleDb.OleDbConnection
    Public CurUser当前用户 As Stru用户信息
    Dim usrName As String
    Dim pwd As String
    Dim WithEvents t1 As New Timer
    Dim UsrStatus用户状态 As enumstatus用户状态
    Public 当前用户剩余时间 As Double

    Dim feerate当前软件费率 As Double
    Dim AppName As String
    Public Event 用户状态改变了(ByRef e As enumstatus用户状态)
    Public Event 用户余额变化了()
    Public Event 用户没有购买本软件()
    Public Event 查无此用户()
    Public Event 有此用户()
    Public Event 此用户为过期或欠费用户()

    Public Sub New(inusrname As String, inpwd As String, infeerate As Double, InAppName As String)
        Try
            usrName = inusrname
            pwd = inpwd
            feerate当前软件费率 = infeerate
            AppName = InAppName
        Catch ex As Exception

        End Try
    End Sub


    Public Sub 连接_构造用户_查询该用户余额并引发欠费事件()
        Try
            AdodbConn = New ADODB.Connection
            AdodbConn.ConnectionString = "DRIVER={MySQL ODBC 8.0 Unicode Driver};" & "SERVER=211.149.234.64;" & "Port=3306;" & "DATABASE=mydata;" & "UID='root';PWD=31415926; OPTION=131072"
            AdodbConn.Open()
            If AdodbConn.State = 1 Then
                With CurUser当前用户
                    .username = usrName
                End With
                查询用户身份和权限清单()
                If CurUser当前用户.用户身份 = Stru用户身份.不存在的用户 Or CurUser当前用户.用户身份 = Stru用户身份.过期用户或欠费用户 Then
                    GoTo 10
                End If
                Dim lst As New List(Of String)
                lst = Split(CurUser当前用户.购买的产品, ",").ToList
                If lst.Exists(Function(ss As String) ss = AppName) = False Then
                    AdodbConn.Close()
                    RaiseEvent 用户没有购买本软件()
                    Exit Sub
                End If
                '得到用户身份后查询用户余额或者断开用户

                sb查询当前用户余额并将余额信息写到当前用户属性并引起用户状态改变事件()
                If 当前用户剩余时间 > 0 Then
                    t1.Interval = 180000
                    t1.Start()
                End If
10:
                '此处连接完毕后关闭==========================================================
                If AdodbConn.State = 1 Then
                    AdodbConn.Close()
                End If
            End If
            'OleDbCon = New OleDb.OleDbConnection("provider=sqloledb;SERVER=211.149.234.64;Port=3306;DATABASE=mydata;User='root';password='31415926'; OPTION=131072")
            'OleDbCon.Open()

        Catch ex As Exception
            MsgBox(Err.Description)
            Try
                AdodbConn.Close()
            Catch ex1 As Exception

            End Try

        End Try

    End Sub

    Sub 扣减用户费用(CheckMoney扣除金额 As Double)
        Try
            Dim cmd As New ADODB.Command
            cmd.CommandText = "update users set money=money-" & CheckMoney扣除金额 & " where username='" & usrName & "'"
            cmd.ActiveConnection = AdodbConn
            cmd.Execute()
        Catch ex As Exception
            MsgBox("更新用户余额出错" & Err.Description)
        End Try


    End Sub

    Sub sb查询当前用户余额并将余额信息写到当前用户属性并引起用户状态改变事件()
        Try
            If AdodbConn.State <> 1 Then
                AdodbConn.Open()
            End If
            If AdodbConn.State = 1 Then
                Dim rs As ADODB.Recordset = New ADODB.Recordset
                rs.CursorLocation = CursorLocationEnum.adUseClient
                rs.Open("SELECT username,userpwd,freenumberoftimes,money,enddatetime,state FROM users where username='" & CurUser当前用户.username & "'", AdodbConn, CursorTypeEnum.adOpenStatic)

                Select Case CurUser当前用户.用户身份
                    Case Stru用户身份.试用用户
                        CurUser当前用户.validDate = rs("enddatetime").Value
                        当前用户剩余时间 = DateDiff(DateInterval.Minute, Now, CDate(CurUser当前用户.validDate))
                        If DateDiff(DateInterval.Minute, Now, CDate(CurUser当前用户.validDate)) < 5 And DateDiff(DateInterval.Minute, Now, CDate(CurUser当前用户.validDate)) > 0 Then

                            If UsrStatus用户状态 <> enumstatus用户状态.费用低 Then
                                Dim e As enumstatus用户状态 = enumstatus用户状态.费用低
                                UsrStatus用户状态 = e
                                RaiseEvent 用户状态改变了(e)
                            End If

                        ElseIf DateDiff(DateInterval.Minute, Now, CDate(CurUser当前用户.validDate)) <= 0 Then
                            If UsrStatus用户状态 <> enumstatus用户状态.欠费 Then
                                Dim e As enumstatus用户状态 = enumstatus用户状态.欠费
                                UsrStatus用户状态 = e
                                RaiseEvent 用户状态改变了(e)
                                RaiseEvent 此用户为过期或欠费用户()
                            End If
                        Else
                            If UsrStatus用户状态 <> enumstatus用户状态.正常 Then
                                Dim e As enumstatus用户状态 = enumstatus用户状态.正常
                                UsrStatus用户状态 = e
                                RaiseEvent 用户状态改变了(e)
                            End If
                        End If

                    Case Stru用户身份.正式用户
                        CurUser当前用户.money = rs("money").Value
                        当前用户剩余时间 = rs("money").Value / feerate当前软件费率
                        If rs("money").Value < 5 And rs("money").Value > 0 Then
                            If UsrStatus用户状态 <> enumstatus用户状态.费用低 Then
                                Dim e As enumstatus用户状态 = enumstatus用户状态.费用低
                                UsrStatus用户状态 = e
                                RaiseEvent 用户状态改变了(e)
                            End If

                        ElseIf CurUser当前用户.money <= 0 Then
                            If UsrStatus用户状态 <> enumstatus用户状态.欠费 Then
                                Dim e As enumstatus用户状态 = enumstatus用户状态.欠费
                                UsrStatus用户状态 = e
                                RaiseEvent 用户状态改变了(e)
                                RaiseEvent 此用户为过期或欠费用户()
                            End If
                        Else
                            If UsrStatus用户状态 <> enumstatus用户状态.正常 Then
                                Dim e As enumstatus用户状态 = enumstatus用户状态.正常
                                UsrStatus用户状态 = e
                                RaiseEvent 用户状态改变了(e)
                            End If
                        End If

                    Case Stru用户身份.管理员
                    Case Stru用户身份.不存在的用户
                        If UsrStatus用户状态 <> enumstatus用户状态.用户不存在 Then
                            Dim e As enumstatus用户状态 = enumstatus用户状态.用户不存在
                            UsrStatus用户状态 = e
                            RaiseEvent 用户状态改变了(e)
                        End If
                        当前用户剩余时间 = 0
                End Select
                If 当前用户剩余时间 < 0 Then
                    t1.Stop()

                End If
                If AdodbConn.State = 1 Then
                    AdodbConn.Close()
                End If
            End If

        Catch ex As Exception
            t1.Stop()
            If AdodbConn.State = 1 Then
                AdodbConn.Close()
            End If
            MsgBox("查询和获取用户信息过程出错" & Err.Description)
        End Try

    End Sub
    Sub 查询用户身份和权限清单()
        Try
            Dim rs As ADODB.Recordset = New ADODB.Recordset
            rs.CursorLocation = CursorLocationEnum.adUseClient
            rs.Open("SELECT username,userpwd,freenumberoftimes,money,enddatetime,state,purchasedapps FROM users where username='" & usrName & "' and userpwd='" & GetMD5FromString(pwd) & "'", AdodbConn, CursorTypeEnum.adOpenStatic)

            If rs.RecordCount < 1 Then
                rs.Close()
                ' System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)
                RaiseEvent 查无此用户()
                With CurUser当前用户
                    .用户身份 = Stru用户身份.不存在的用户
                End With
                Exit Sub
            Else
                RaiseEvent 有此用户()
            End If

            If IsDBNull(rs("state").Value) Then
                rs.Close()
                'System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)
                With CurUser当前用户
                    .用户身份 = Stru用户身份.过期用户或欠费用户
                    RaiseEvent 此用户为过期或欠费用户()
                End With

            Else
                If rs("state").Value = 3 Then
                    rs.Close()
                    'System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)
                    With CurUser当前用户
                        .用户身份 = Stru用户身份.管理员
                    End With
                    Exit Sub
                End If
                If rs("state").Value = 2 Then

                    ' System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)
                    With CurUser当前用户
                        .用户身份 = Stru用户身份.试用用户
                        If IsDBNull(rs("purchasedapps").Value) = False Then
                            .购买的产品 = rs("purchasedapps").Value
                        End If
                    End With
                    rs.Close()
                    Exit Sub
                End If

                If rs("state").Value = 0 Then

                    'System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)
                    With CurUser当前用户
                        .用户身份 = Stru用户身份.正式用户
                        If IsDBNull(rs("purchasedapps").Value) = False Then
                            .购买的产品 = rs("purchasedapps").Value
                        End If
                    End With
                    rs.Close()
                    Exit Sub
                End If
            End If

        Catch ex As Exception
            MsgBox("查询用户身份过程出错" & Err.Description)

        End Try

    End Sub
    Public Function GetMD5FromString(ByVal msg As String) As String

        '1.创建一个用来计算MD5值的类的对象
        Dim md5 As MD5 = MD5.Create
        'Imports (MD5 md5 = MD5.Create())
        '把字符串转换为byte[]
        '注意：如果字符串中包含汉字，则这里会把汉字使用utf-8编码转换为byte[]，当其他地方
        '计算MD5值的时候，如果对汉字使用了不同的编码，则同样的汉字生成的byte[]是不一样的，所以计算出的MD5值也就不一样了。
        Dim msgBuffer() As Byte = Encoding.Default.GetBytes(msg)

        '2.计算给定字符串的MD5值
        '返回值就是就算后的MD5值,如何把一个长度为16的byte[]数组转换为一个长度为32的字符串：就是把每个byte转成16进制同时保留2位即可。
        Dim md5Buffer() As Byte = md5.ComputeHash(msgBuffer)
        md5.Clear() '释放资源

        Dim sbMd5 As StringBuilder = New StringBuilder()
        Dim i As Integer
        For i = 0 To md5Buffer.Length - 1 Step i + 1
            sbMd5.Append(md5Buffer(i).ToString("x2"))
        Next
        Return sbMd5.ToString()
    End Function

    Private Sub t1_Tick(sender As Object, e As EventArgs) Handles t1.Tick
        Try
            If AdodbConn.State <> 1 Then
                AdodbConn.Open()
            End If
            If CurUser当前用户.用户身份 = Stru用户身份.正式用户 Then
                扣减用户费用(t1.Interval / 60000 * feerate当前软件费率)
            End If
            sb查询当前用户余额并将余额信息写到当前用户属性并引起用户状态改变事件()
            RaiseEvent 用户余额变化了()
            If AdodbConn.State = 1 Then AdodbConn.Close()
        Catch ex As Exception
            MsgBox("时钟事件出错" & Err.Description)
            If AdodbConn.State = 1 Then AdodbConn.Close()
        End Try
    End Sub
End Class
