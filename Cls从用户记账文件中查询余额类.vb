Imports System.Data
Imports System.Data.Common
Imports System.Collections.Generic
Imports System.Text
Imports ADODB
Imports System.IO
Imports System.Security.Cryptography
Public Structure Stru当前用户
    Dim username As String
    Dim 用户身份 As UserIdentity
    Dim freeCount As Integer
    Dim 剩余时间 As Double

    Dim validDate As String

End Structure
Public Enum UserIdentity
    管理员 = 1
    试用用户 = 2
    正式用户 = 3
    不存在的用户 = 4
    过期用户或欠费用户 = 5
End Enum



Public Class Cls从用户记账文件中查询余额类


    Public 当前用户 As Stru当前用户
    Public Event RaiseAlert()
    Public Event UserTimeOut()
    Public Event UserMoneyOut()
    Public 用户记账文件 As String



    'Public rs As ADODB.Recordset
    Public Sub New()

    End Sub





    Function ReadUserInfo() As Boolean
        Try


            Dim s() As String = File.ReadAllLines(用户记账文件)
            With 当前用户
                .username = s(0)
                If s(1) = "用户类型：试用用户" Then .用户身份 = UserIdentity.试用用户
                If s(1) = "用户类型：正式用户" Then .用户身份 = UserIdentity.正式用户
                If s(1) = "用户类型：管理员" Then .用户身份 = UserIdentity.管理员
                If s(1) = "用户类型：不存在的用户" Then .用户身份 = UserIdentity.不存在的用户
                If s(1) = "用户类型：过期用户或欠费用户" Then .用户身份 = UserIdentity.过期用户或欠费用户
                If .用户身份 = UserIdentity.试用用户 Then .validDate = s(2)
                If .用户身份 = UserIdentity.正式用户 Then .剩余时间 = s(2)

            End With




            Select Case 当前用户.用户身份
                Case UserIdentity.试用用户

                    If DateDiff(DateInterval.Minute, Now, CDate(当前用户.validDate)) < 5 And DateDiff(DateInterval.Minute, Now, CDate(当前用户.validDate)) > 0 Then
                        RaiseEvent RaiseAlert()
                    ElseIf DateDiff(DateInterval.Minute, Now, CDate(当前用户.validDate)) <= 0 Then
                        RaiseEvent UserTimeOut()
                    End If
                Case UserIdentity.正式用户

                    If 当前用户.剩余时间 < 5 And 当前用户.剩余时间 > 0 Then
                        RaiseEvent RaiseAlert()
                    ElseIf 当前用户.剩余时间 <= 0 Then
                        RaiseEvent UserMoneyOut()

                    End If
            End Select
            Return True
        Catch ex As Exception
            MsgBox("读取当前用户信息过程出错" & Err.Description)
            Return False
        End Try
    End Function


End Class
